///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Event Log Information Class Implementation
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

// Since Vista the Event ids have changed in the Security Log, effectively
// there is now an offset of +4096, important ones are      Pre-Vista   Vista
//      Windows is starting up.                                         4608
//      The system time was changed.                                    4616
//      An account was successfully logged on               512         4624
//      An account was logged off.                          538         4634
//      User initiated logoff:                              551         4647
//      A logon was attempted using explicit credentials.   551         4648
//      Special privileges assigned to new logon.           552         4672
//      A new process has been created                      592         4688
//      A process has exited                                593         4689
//      A primary token was assigned to process.            600         4696
//
//  Remarks:
//      Logon Types
//      2   - Interactive
//      3   - Network
//      4   - Batch
//      5   - Service
//      6   - Proxy
//      7   - Unlock Workstation
//      8   - NetworkCleartext'
//      9   - NewCredentials'
//      10  - RemoteInteractive'
//      11  - CachedInteractive'
//      13  - CachedRemoteInteractive'
//      14  - CachedUnlock'
//      (0 & 1 are invalid)
//
//      As seen in event logs the the machine account i.e. computer_name$ uses
//      logon type of 0 which according to the above list in the documentation
//      is invalid
//
//      Event log is on NT only, this uses the security log
//      Windows NT is starting up.              Event ID 512
//      Windows NT is shutting down.            Event ID 513
//      Interactive logon                       Event ID 528        Type 2
//         Network logon                           Event ID 528     Type 3
//         Net Use connection                      Event ID 528     Type 3
//         Service logon                           Event ID 528     Type 5
//      Interactive logoff                      Event ID 538        Type 2
//         Network logoff                          Event ID 538     Type 3
//         Net use disconnection                   Event ID 538     Type 3
//         Autodisconnect                          Event ID 538     Type 3
//      Successful Network Logon                Event ID 540        Type 3
//      User initiated logoff                   Event ID 551
//      A new process has been created          Event ID 592
//      A process has exited:                   Event ID 593
//        Session reconnected to winstation        Event ID 682
//        Session disconnected from winstation     Event ID 683
//
//      When logon with 540 the corresponding logoff is 538 type 3
//      If no log off is found for a logon, then will ignore it for the total
//      time logged on.
//
//      Despite the documentation, NT3.51, 4.0 and 2000 log system
//      shutdowns = 513
//
//      For XP/2003 the format for Event ID 593 is:
//          A process has exited:
//              Process ID:     %1
//              Image File Name:%2  (full path)
//              User Name:      %3
//              Domain:         %4
//              Logon ID:       %5
//
//      For NT3.51/NT4/2000 the format for Event ID 593 is:
//          A process has exited
//              Process ID:     %1
//              User Name:      %2
//              Domain:         %3
//              Logon ID:       %4
//
//      Services:
//          Event ID 592 with type 5 seems to be executables that were started
//          BY a service, not the actual service itself. So its not very useful
//          for counting executables started as services. Typically these are
//          started under the NT AUTHORITY\SYSTEM account so do not count
//          services separately.
//
//      Logon IDs:
//          (0x0,0x3E4) =   User Name:    NETWORK SERVICE Domain: NT AUTHORITY
//          (0x0,0x3E5) =   User Name:    LOCAL SERVICE   Domain: NT AUTHORITY
//          (0x0,0x3E7) =   User Name:    SYSTEM          Domain: NT AUTHORITY
//
//      Attempts to count the number of seconds each process was running was
//      not very successful because need to maintain a lengthy list of process
//      starts and if the audit policy is just set to record only/mostly
//      process event there can be many thousands. Also, process exit events
//      are not always reliably recorded so the duration sometime erroneous.
//
///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/EventLogInformation.h"

// 2. C System Files
#include <time.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Directory.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/FileVersion.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/Registry.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/ProcessInformation.h"
#include "WinAudit/Header Files/WindowsInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
EventLogInformation::EventLogInformation()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
EventLogInformation::~EventLogInformation()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

//
// Assignment operator - not allowed so no implementation
//

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Fill an array of audit failures found in the security event log
//
//  Parameters:
//      maximumEntries - maximum number of entries to return
//      pRecords       - receives the audit records
//
//  Returns:
//      void
//===============================================================================================//
void EventLogInformation::GetAuditFailureRecords(
                                            size_t maximumEntries,
                                            TArray< AuditRecord >* pRecords )
{
    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    GetSourceNameAuditRecords( PXS_CATEGORY_SECURITY_LOG,
                               PXS_SECURITY_LOG_TIME_GENERATED,
                               PXS_SECURITY_LOG_FILE,
                               PXS_SECURITY_LOG_SOURCE_NAME,
                               PXS_SECURITY_LOG_DESCRIPTION,
                               L"Security",
                               EVENTLOG_AUDIT_FAILURE,
                               maximumEntries,
                               pRecords );
}

//===============================================================================================//
//  Description:
//        Get event log errors
//
//  Parameters:
//      maximumEntries - maximum number of entries to return
//      pRecords       - receives the audit records
//
//  Returns:
//      void
//===============================================================================================//
void EventLogInformation::GetErrorRecords( size_t maximumEntries, TArray< AuditRecord >* pRecords )
{
    size_t i = 0, numNames = 0;
    StringArray SourceNames;
    TArray< AuditRecord > SourceNameRecords;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    GetSourceNames( &SourceNames );
    SourceNames.Sort( true );
    numNames = SourceNames.GetSize();
    while ( ( i < numNames ) && ( pRecords->GetSize() < maximumEntries ) )
    {
        SourceNameRecords.RemoveAll();
        GetSourceNameAuditRecords( PXS_CATEGORY_ERROR_LOG,
                                   PXS_ERROR_LOG_GENERATED,
                                   PXS_ERROR_LOG_FILE,
                                   PXS_ERROR_LOG_SOURCE_NAME,
                                   PXS_ERROR_LOG_DESCRIPTION,
                                   SourceNames.Get( i ),
                                   EVENTLOG_ERROR_TYPE,
                                   maximumEntries - pRecords->GetSize(),
                                   &SourceNameRecords );
        pRecords->Append( SourceNameRecords );
        i++;
    }

    // Keep in limit
    if ( pRecords->GetSize() > maximumEntries )
    {
        pRecords->SetSize( maximumEntries );
    }
}

//===============================================================================================//
//  Description:
//      Get software/executable usage data as an array of audit records
//
//  Parameters:
//      pRecords - audit record array to receive the data
//
//  Returns:
//      void
//===============================================================================================//
void EventLogInformation::GetSoftwareMeteringRecords( TArray< AuditRecord >* pRecords )
{
    String      FileName, Publisher, VersionString, Description;
    String      FilePath, FirstStart, LastStart, NumStarts;
    Formatter   Format;
    AuditRecord Record;
    TYPE_METERING_DATA  Element;
    WindowsInformation  WindowsInfo;
    TList< EventLogRecord >     EventLogRecords;
    TList< TYPE_METERING_DATA > MeteringData;
    const EventLogRecord*       pEventLogRecord = nullptr;
    const TYPE_METERING_DATA*   pMD = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    // Event log identifier changed with Vista. The string number is the same
    // NT 3-5 uses "Image File Name", NT6 uses "NewProcessName"
    DWORD EVENT_ID_PROCESS_CREATED = 592;
    DWORD STRING_NUMBER_FILENAME   = 2;
    if ( WindowsInfo.GetMajorVersion() >= 6 )
    {
        EVENT_ID_PROCESS_CREATED = 592 + 4096;
    }

    // Get the process created records from the security log
    DWORD EventIDs[] = { EVENT_ID_PROCESS_CREATED };
    GetEventLogRecords( L"Security",
                        true,               // EVENTLOG_FORWARDS_READ
                        EventIDs,
                        ARRAYSIZE( EventIDs ),
                        0,                  // Not used
                        SIZE_MAX,           // No limit
                        &EventLogRecords );
    if ( EventLogRecords.IsEmpty() )
    {
        return;
    }

    EventLogRecords.Rewind();
    do
    {
        pEventLogRecord = EventLogRecords.GetPointer();
        pEventLogRecord->GetString( STRING_NUMBER_FILENAME, &FileName );
        FileName.Trim();
        memset( &Element, 0, sizeof ( Element ) );
        if ( FindMeteringData( FileName, &MeteringData, &Element ) )
        {
            // Update exisiting element
            Element.lastStartTime = PXSMaxTimeT( pEventLogRecord->GetTimeGenerated(),
                                                 Element.lastStartTime );
            Element.numberStarts++;
            MeteringData.Set( Element );
        }
        else
        {
            // Append a new element to the list
            Element.firstStartTime = pEventLogRecord->GetTimeGenerated();
            Element.lastStartTime  = pEventLogRecord->GetTimeGenerated();
            Element.numberStarts   = 1;
            StringCchCopy( Element.szFileName, ARRAYSIZE(Element.szFileName ), FileName.c_str() );
            MeteringData.Append( Element );
        }
    } while ( EventLogRecords.Advance() );

    // Make the audit records
    if ( MeteringData.IsEmpty() )
    {
        return;     // Nothing more to do
    }

    MeteringData.Rewind();
    do
    {
        pMD = MeteringData.GetPointer();
        FilePath     = pMD->szFileName;
        FileName     = PXS_STRING_EMPTY;
        Publisher    = PXS_STRING_EMPTY;
        VersionString= PXS_STRING_EMPTY;
        Description  = PXS_STRING_EMPTY;
        GetFileNameAndVersionInfo( FilePath, &FileName, &Publisher, &VersionString, &Description );
        FirstStart     = Format.TimeTToLocalTimeInIso( pMD->firstStartTime );
        LastStart      = Format.TimeTToLocalTimeInIso( pMD->lastStartTime );
        NumStarts      = Format.UInt32( pMD->numberStarts );
        Record.Reset( PXS_CATEGORY_SOFTWARE_METERING );
        Record.Add( PXS_SOFT_METER_FILE_NAME     , FileName );
        Record.Add( PXS_SOFT_METER_FILE_PATH     , FilePath );
        Record.Add( PXS_SOFT_METER_FILE_VERSION  , VersionString );
        Record.Add( PXS_SOFT_METER_PUBLISHER     , Publisher );
        Record.Add( PXS_SOFT_METER_FIRST_START_TS, FirstStart);
        Record.Add( PXS_SOFT_METER_LAST_START_TS , LastStart );
        Record.Add( PXS_SOFT_METER_NUMBER_STARTS , NumStarts );
        Record.Add( PXS_SOFT_METER_REMOTE_STARTS , PXS_STRING_EMPTY );
        Record.Add( PXS_SOFT_METER_OTHER_STARTS  , PXS_STRING_EMPTY );
        pRecords->Add( Record );
    } while ( MeteringData.Advance() );

    PXSSortAuditRecords( pRecords, PXS_SOFT_METER_FILE_NAME );
}

//===============================================================================================//
//  Description:
//      Get the event log file source names
//
//  Parameters:
//      pSourceNames - string array to receive source names
//
//  Remarks:
//      NT4 has Application, Security and System, later operating systems have
//      custom log files such as Directory Service
//
//  Returns:
//      void
//===============================================================================================//
void EventLogInformation::GetSourceNames(  StringArray* pSourceNames )
{
    Registry RegObject;

    if ( pSourceNames == nullptr )
    {
        throw ParameterException( L"pSourceNames", __FUNCTION__ );
    }
    pSourceNames->RemoveAll();

    RegObject.Connect( HKEY_LOCAL_MACHINE );
    RegObject.GetSubKeyList( L"System\\CurrentControlSet\\Services\\EventLog", pSourceNames );
    pSourceNames->Sort( true );
}

//===============================================================================================//
//  Description:
//      Get the uptime statistics from the event log
//
//  Parameters:
//      pRecord - receives the audit record
//
//  Remarks:
//      MDSN Q196452
//      NT4 with service pack 4
//
//      Event 6005 is logged at boot time noting that the Event Log service
//      was started. It gives the message "The Event log service was started".
//
//      Event 6006 is logged as a clean shutdown. It gives the message "The
//      Event log service was stopped".
//
//      Event 6008 is logged as a dirty shutdown. It gives the message "The
//      previous system shutdown at time on date was unexpected".
//
//      Event 6009 is logged during every boot and indicates the operating
//      system version, build number, service pack level, and other pertinent
//      information about the system. Depending on your current configuration,
//      it gives a message similar to: "Microsoft (R) Windows NT 4.0 1381
//      Service Pack 6 Multiprocessor free".
//
//  Returns:
//      void
//===============================================================================================//
void EventLogInformation::GetUptimeStatsRecord( AuditRecord* pRecord )
{
    WORD      eventID = 0;
    DWORD     timesBooted = 0, cleanShutDowns = 0, unexpectedShutDowns = 0;
    DWORD     eventCounter = 0;
    double    availability = 0.0;
    time_t    timeNow = 0, timeGenerated = 0, startDate = 0, systemUptime = 0;
    time_t    lastTimeGenerated = 0, bootTime = 0, totalUpTime = 0;
    String    Value;
    Formatter Format;
    SystemInformation SystemInfo;
    TList<EventLogRecord> EventLogRecords;
    const EventLogRecord* pEventLogRecord = nullptr;

    if ( pRecord == nullptr )
    {
        throw ParameterException( L"pRecord", __FUNCTION__ );
    }
    pRecord->Reset( PXS_CATEGORY_UPTIME );

    // Get the events from the security log
    DWORD EventIDs[] = { 6005, 6006, 6008, 6009 };
    GetEventLogRecords( L"System",
                        true,              // EVENTLOG_FORWARDS_READ
                        EventIDs,
                        ARRAYSIZE( EventIDs ),
                        0,
                        SIZE_MAX,           // All records
                        &EventLogRecords );
    if ( EventLogRecords.IsEmpty() )
    {
        return;     // Nothing more to do
    }

    timeNow = time( nullptr );
    EventLogRecords.Rewind();
    do
    {
        pEventLogRecord = EventLogRecords.GetPointer();
        eventCounter = PXSAddUInt32( eventCounter, 1 );

        // For the first event, store the date, the form is in seconds
        // elapsed since 00:00:00 January 1, 1970, UTC
        timeGenerated = pEventLogRecord->GetTimeGenerated();
        if ( eventCounter == 1 )
        {
            startDate = timeGenerated;
        }

        // Ensure items that are in chronological order or after the
        // current time.
        if ( ( timeGenerated <  timeNow ) &&
             ( timeGenerated >= lastTimeGenerated ) )
        {
            lastTimeGenerated = timeGenerated;
            eventID = pEventLogRecord->GetEventID();
            if ( eventID == 6009 )  // System boot ups
            {
                timesBooted = PXSAddUInt32( timesBooted, 1 );
                bootTime = timeGenerated;       // UTC
            }
            else if ( eventID == 6006 )         // Clean shut downs
            {
                cleanShutDowns = PXSAddUInt32( cleanShutDowns, 1 );
            }
            else if ( eventID == 6008 )     // Dirty shut downs
            {
                unexpectedShutDowns = PXSAddUInt32( unexpectedShutDowns, 1 );
            }

            // Sum up the total uptime if the event was a shut down
            if ( ( bootTime      > 0 ) &&
                 ( timeGenerated > bootTime ) &&
                 ( ( eventID == 6006 ) || ( eventID == 6008 ) ) )
            {
                totalUpTime += ( timeGenerated - bootTime );
                bootTime     = 0;   // Reset
            }
        }
    } while ( EventLogRecords.Advance() );

    // Make the audit record, check the start date, if none, then no data
    if ( ( startDate > 0 ) && ( timeNow > startDate ) )
    {
        // Start of data, seconds since 1/1/70 00:00:00
        Value = Format.TimeTToLocalTimeInIso( startDate );
        pRecord->Add( PXS_UPTIME_DATA_START, Value );

        // System up time since the last boot
        Value        = PXS_STRING_EMPTY;
        systemUptime = SystemInfo.GetUpTime();
        if ( systemUptime )
        {
            Value = Format.SecondsToDDHHMM( systemUptime );
        }
        pRecord->Add( PXS_UPTIME_SYSTEM_UPTIME, Value );

        // System availability
        availability = ( 100.0 * totalUpTime ) / ( timeNow - startDate );
        Value        = Format.Double( availability, 2 );
        Value       += L"%";
        pRecord->Add( PXS_UPTIME_SYSTEM_AVAIL_PERCENT, Value );

        // Total Uptime
        Value = PXS_STRING_EMPTY;
        if ( totalUpTime > 0 )
        {
            Value = Format.SecondsToDDHHMM( totalUpTime );
        }
        pRecord->Add( PXS_UPTIME_TOTAL_UPTIME, Value );

        // Total Downtime
        Value = PXS_STRING_EMPTY;
        if ( timeNow > ( startDate + totalUpTime ) )
        {
            Value = Format.SecondsToDDHHMM( timeNow - (startDate+totalUpTime) );
        }
        pRecord->Add( PXS_UPTIME_TOTAL_DOWNTIME, Value );

        // Boot data
        Value = Format.UInt32( timesBooted );
        pRecord->Add( PXS_UPTIME_TIMES_BOOTED,  Value );

        Value = Format.UInt32( cleanShutDowns );
        pRecord->Add( PXS_UPTIME_CLEAN_SHUTDOWNS,  Value );

        Value = Format.UInt32( unexpectedShutDowns );
        pRecord->Add( PXS_UPTIME_UNEXPECTED_SHUTDOWNS, Value );
    }
}

//===============================================================================================//
//  Description:
//      Get user logon/logoff events
//
//  Parameters:
//      pRecords - array to receive the audit records
//
//  Returns:
//      void
//===============================================================================================//
void EventLogInformation::GetUserLogonRecords( TArray< AuditRecord >* pRecords )
{
    time_t      timeGenerated = 0;
    String      DomainUserName, Username, Domain, LogonType, Value;
    Formatter   Format;
    AuditRecord Record;
    TYPE_USER_DATA     Element;
    WindowsInformation WindowsInfo;
    TList<TYPE_USER_DATA> UserData;
    TList<EventLogRecord> EventLogRecords;
    const TYPE_USER_DATA* pUserData       = nullptr;
    const EventLogRecord* pEventLogRecord = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

     // If on Vista or newer, security event messages have changed
    DWORD EVENT_ID_LOGON         = 528;         // Logon
    DWORD STRING_NUM_USER_NAME   = 1;           // NT 3-5 "User name"
    DWORD STRING_NUM_DOMAIN_NAME = 2;           // NT 3-5 "Domain name"
    DWORD STRING_NUM_LOGON_TYPE  = 4;           // NT 3-5 "Logon Type"
    if ( WindowsInfo.GetMajorVersion() >= 6 )
    {
        // Event IDs are offset by 4096
        EVENT_ID_LOGON         = 528 + 4096;    // NT6+ Logon
        STRING_NUM_USER_NAME   = 6;             // NT6+ "TargetUserName"
        STRING_NUM_DOMAIN_NAME = 7;             // NT6+ "TargetDomainName"
        STRING_NUM_LOGON_TYPE  = 9;             // NT6+ "LogonType"
    }

    // Get the records from the Security log
    DWORD EventIDs[] = { EVENT_ID_LOGON };
    GetEventLogRecords( L"Security",
                        true,               // EVENTLOG_FORWARDS_READ
                        EventIDs,
                        ARRAYSIZE( EventIDs ),
                        0,
                        SIZE_MAX,           // No limit
                        &EventLogRecords );
    if ( EventLogRecords.IsEmpty() )
    {
        return;     // Nothing more to do
    }

    EventLogRecords.Rewind();
    do
    {
        pEventLogRecord = EventLogRecords.GetPointer();
        Username  = PXS_STRING_EMPTY;
        Domain    = PXS_STRING_EMPTY;
        LogonType = PXS_STRING_EMPTY;
        pEventLogRecord->GetString( STRING_NUM_USER_NAME  , &Username );
        pEventLogRecord->GetString( STRING_NUM_DOMAIN_NAME, &Domain   );
        pEventLogRecord->GetString( STRING_NUM_LOGON_TYPE , &LogonType);

        // See if the user name is already in the list. Express the user
        // as domain\user_name. If no match add a new item to the list
        DomainUserName = Domain;
        if ( DomainUserName.GetLength() )
        {
            DomainUserName += PXS_PATH_SEPARATOR;
        }
        DomainUserName += Username;
        DomainUserName.Trim();
        if ( DomainUserName.GetLength() )
        {
            timeGenerated = pEventLogRecord->GetTimeGenerated();
            memset( &Element, 0, sizeof ( Element ) );
            if ( FindUserData( DomainUserName, &UserData, &Element ) )
            {
                // Update exisiting element
                Element.lastLogonTime = PXSMaxTimeT( timeGenerated,
                                                     Element.lastLogonTime );
                if ( LogonType.CompareI( L"2" ) == 0 )
                {
                    Element.consoleLogons++;
                }
                else if ( LogonType.CompareI( L"10" ) == 0 )
                {
                    Element.remoteLogons++;
                }
                else
                {
                    Element.otherLogons++;
                }
                UserData.Set( Element );
            }
            else
            {
                // Append a new element
                Element.firstLogonTime = timeGenerated;
                Element.lastLogonTime  = timeGenerated;
                if ( LogonType.CompareI( L"2" ) == 0 )
                {
                    Element.consoleLogons = 1;
                }
                else if ( LogonType.CompareI( L"10" ) == 0 )
                {
                    Element.remoteLogons = 1;
                }
                else
                {
                    Element.otherLogons = 1;
                }
                StringCchCopy( Element.szUserName,
                               ARRAYSIZE( Element.szUserName ), DomainUserName.c_str() );
                UserData.Append( Element );
            }
        }
    } while ( EventLogRecords.Advance() );

    // Make the audit record
    if ( UserData.IsEmpty() )
    {
        return;     // Nothing more to do
    }

    UserData.Rewind();
    do
    {
        pUserData = UserData.GetPointer();
        Record.Reset( PXS_CATEGORY_USER_LOGONS );

        Value = pUserData->szUserName;
        Record.Add( PXS_USER_LOGONS_USER_NAME, Value );

        Value = Format.TimeTToLocalTimeInIso( pUserData->firstLogonTime );
        Record.Add( PXS_USER_LOGONS_FIRST_LOGON_TS, Value );

        Value = Format.TimeTToLocalTimeInIso( pUserData->lastLogonTime );
        Record.Add( PXS_USER_LOGONS_LAST_LOGON_TS, Value );

        Value = Format.UInt32( pUserData->consoleLogons );
        Record.Add( PXS_USER_LOGONS_CONSOLE_LOGONS, Value );

        Value = Format.UInt32( pUserData->remoteLogons );
        Record.Add( PXS_USER_LOGONS_REMOTE_LOGONS, Value );

        Value = Format.UInt32( pUserData->otherLogons );
        Record.Add( PXS_USER_LOGONS_OTHER_LOGONS, Value );

        pRecords->Add( Record );
    } while ( UserData.Advance() );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Find the element in the software metering list matching the specified
//      file name. If the element is found, the list is positioned on that
//      node. If the element is not found the list is positioned at the tail.
//
//  Parameters:
//      pMeteringData  - the file name to search for
//      MeteringData   - the list of software metering events
//      pElement       - receives the element if it is found
//
//  Returns:
//      true if the element was found otherwise false
//===============================================================================================//
bool EventLogInformation::FindMeteringData( const String& FileName,
                                            TList< TYPE_METERING_DATA>* pMeteringData,
                                            TYPE_METERING_DATA* pElement )
{
    bool  found = false;
    const TYPE_METERING_DATA* pData  = nullptr;

    if ( pElement == nullptr )
    {
        throw ParameterException( L"pElement", __FUNCTION__ );
    }

    if ( ( pMeteringData == nullptr ) ||
         ( pMeteringData->IsEmpty() ) ||
         ( FileName.IsEmpty()       )  )
    {
        return false;
    }

    pMeteringData->Rewind();
    do
    {
        pData = pMeteringData->GetPointer();
        if ( FileName.CompareI( pData->szFileName ) == 0 )
        {
            memcpy( pElement, pData, sizeof ( TYPE_METERING_DATA ) );
            found = true;
        }
    } while ( ( found == false ) && ( pMeteringData->Advance() ) );

    return found;
}

//===============================================================================================//
//  Description:
//      Find the element in the user data list by matching on the specified
//      user name. If the element is found, the list is positioned on that
//      node. If the element is not found the list is positioned at the tail.
//
//  Parameters:
//      DomainUserName  - the user name to search for, form is
//      pUserData       - the list of user data events
//                          DomainName\UserName
//      pElement        - receives the element if it is found
//
//  Returns:
//      true if the element was found otherwise false
//===============================================================================================//
bool EventLogInformation::FindUserData( const String& DomainUserName,
                                        TList< TYPE_USER_DATA >* pUserData,
                                        TYPE_USER_DATA* pElement )
{
    bool  found = false;
    const TYPE_USER_DATA* pData = nullptr;

    if ( pElement == nullptr )
    {
        throw ParameterException( L"pElement", __FUNCTION__ );
    }

    if ( ( pUserData == nullptr     ) ||
         ( pUserData->IsEmpty()     ) ||
         ( DomainUserName.IsEmpty() )  )
    {
        return false;
    }

    pUserData->Rewind();
    do
    {
        pData = pUserData->GetPointer();
        if ( DomainUserName.CompareI( pData->szUserName ) == 0 )
        {
            memcpy( pElement, pData, sizeof ( TYPE_USER_DATA ) );
            found = true;
        }
    } while ( ( found == false ) && ( pUserData->Advance() ) );

    return found;
}

//===============================================================================================//
//  Description:
//      Get audit events for the event log
//
//  Parameters:
//      auditCategoryID     - audit category identifier
//      timeGeneratedItemID - item id of the time generated field
//      sourceItemID        - item id of the source
//      dwDecriptionItemID  - item id of the description
//      pszLogFile          - the name of the log file e.g. application
//      eventType           - event type, e.g. EVENTLOG_ERROR_TYPE
//      maximumEntries      - maximum number of entries to return
//      pRecords            - receives the data
//
//  Remarks:
//      Often event log error descriptions are repeated, so will suppress
//      repetitions.
//
//  Returns:
//      void
//===============================================================================================//
void EventLogInformation::GetSourceNameAuditRecords( DWORD auditCategoryID,
                                                     DWORD timeGeneratedItemID,
                                                     DWORD logFileItemID,
                                                     DWORD sourceItemID,
                                                     DWORD dwDecriptionItemID,
                                                     LPCWSTR pszLogFile,
                                                     WORD  eventType,
                                                     size_t maximumEntries,
                                                     TArray< AuditRecord >* pRecords )
{
    time_t      timeT = 0;
    String      Description, TimeGenerated, SourceName;
    Formatter   Format;
    AuditRecord Record;
    StringArray Descriptions;
    TList<EventLogRecord> EventLogRecords;
    const EventLogRecord* pEventLogRecord = nullptr;

    if ( pRecords == nullptr )
    {
        throw ParameterException( L"pRecords", __FUNCTION__ );
    }
    pRecords->RemoveAll();

    if ( pszLogFile == nullptr )
    {
        return;     // Nothing to do
    }

    if ( timeGeneratedItemID == 0 )
    {
        throw ParameterException( L"timeGeneratedItemID", __FUNCTION__);
    }

    if ( sourceItemID == 0 )
    {
        throw ParameterException( L"sourceItemID", __FUNCTION__);
    }

    if ( dwDecriptionItemID == 0 )
    {
        throw ParameterException( L"dwDecriptionItemID", __FUNCTION__);
    }

    GetEventLogRecords( pszLogFile,
                        false,         // = EVENTLOG_BACKWARDS_READ
                        nullptr,
                        0, eventType, maximumEntries, &EventLogRecords );
    if ( EventLogRecords.IsEmpty() )
    {
        return;     // Nothing more to do
    }

    EventLogRecords.Rewind();
    do
    {
        pEventLogRecord = EventLogRecords.GetPointer();
        Description = PXS_STRING_EMPTY;
        MakeEventLogDescription( pszLogFile, pEventLogRecord, &Description );

        // Keep an array of event log descriptions as will only record
        // unique descriptions
        if ( Descriptions.AddUniqueI( Description.c_str() ) )
        {
            Record.Reset( auditCategoryID );

            timeT = pEventLogRecord->GetTimeGenerated();
            TimeGenerated = Format.TimeTToLocalTimeInIso( timeT );
            Record.Add( timeGeneratedItemID, TimeGenerated );

            Record.Add( logFileItemID, pszLogFile );

            SourceName = PXS_STRING_EMPTY;
            pEventLogRecord->GetSourceName( &SourceName );
            Record.Add( sourceItemID, SourceName );

            Record.Add( dwDecriptionItemID, Description  );

            pRecords->Add( Record );
        }
    } while ( EventLogRecords.Advance() );
}

//===============================================================================================//
//  Description:
//      Get records from the event log matching the specified event
//      identifiers or type
//
//  Parameters:
//      pszName         - the event log file name
//      forwardsRead    - true to read forward otherwise backwards
//      pEventIDs       - pointer to an array of event ids to return records
//                        for, this parameter can be null
//      numEventIDs     - the number of elements in pEventIDs, use zero if
//                        pEventIDs is nullptr
//      eventType       - the event type
//      maximumRecords  - maximum number of records to return, 0 implies all
//      pEventLogRecords- receives a list of type EVENTLOGRECORD
//
//  Remarks:
//
//  Returns:
//      void
//===============================================================================================//
void EventLogInformation::GetEventLogRecords( LPCWSTR      pszName,
                                              bool         forwardsRead,
                                              const DWORD* pEventIDs,
                                              size_t       numEventIDs,
                                              WORD         eventType,
                                              size_t       maximumRecords,
                                              TList<EventLogRecord>* pEventLogRecords )
{
    const DWORD BYTE_BUF_LEN  = 64 * 1024;   // 64K, should be enough
    bool      privilegeChanged = false;
    WORD      eventID    = 0;
    BYTE*     pBuffer    = nullptr;
    DWORD     readFlags  = 0, bytesRead = 0, idxOffset = 0, minBytesNeeded = 0;
    size_t    numRecords = 0;
    HANDLE    hEventLog  = nullptr;
    String    Insert1;
    Formatter Format;
    AllocateBytes   AllocBytes;
    EventLogRecord  Event;
    EVENTLOGRECORD* pRecord    = nullptr;
    ProcessInformation ProcessInfo;
    WindowsInformation WindowsInfo;

    if ( pEventLogRecords == nullptr )
    {
        throw ParameterException( L"pEventLogRecords", __FUNCTION__ );
    }
    pEventLogRecords->RemoveAll();

    if ( pszName == nullptr )
    {
        throw ParameterException( L"pszName", __FUNCTION__ );
    }

    // Set the maximum number of records to retrieve
    if ( maximumRecords == 0 )
    {
        maximumRecords = DWORD_MAX;     // i.e. all records
    }

    // Try to set SE_SECURITY_NAME (SeSecurityPrivilege)
    try
    {
        if ( ProcessInfo.IsPrivilegeEnabled( SE_SECURITY_NAME ) == false )
        {
            // Do not have SE_SECURITY_NAME privilege, try to acquire it.
            ProcessInfo.ChangePrivilege( SE_SECURITY_NAME, true );
            privilegeChanged = true;
        }
    }
    catch ( const Exception& e )
    {
        // Log it but continue
        PXSLogException( L"Failed to get/alter privilege SE_SECURITY_NAME.", e, __FUNCTION__ );
    }

    // Sequential read
    pBuffer = AllocBytes.New( BYTE_BUF_LEN );
    if ( forwardsRead )
    {
       readFlags = EVENTLOG_SEQUENTIAL_READ | EVENTLOG_FORWARDS_READ;
    }
    else
    {
       readFlags = EVENTLOG_SEQUENTIAL_READ | EVENTLOG_BACKWARDS_READ;
    }

    Insert1 = pszName;
    PXSLogAppInfo1( L"Opening Event Log '%%1'.", Insert1 );
    hEventLog = OpenEventLog( nullptr, pszName );
    if ( hEventLog == nullptr )
    {
        throw SystemException( GetLastError(), L"OpenEventLog", __FUNCTION__ );
    }

    // Catch all exceptions as need to clean up
    try
    {
        // Read the event log
        while ( ReadEventLog( hEventLog,
                              readFlags,
                              0, pBuffer, BYTE_BUF_LEN, &bytesRead, &minBytesNeeded ) )
        {
            // Get each EVENTLOGRECORD from the buffer
            idxOffset = 0;
            pRecord   = reinterpret_cast<EVENTLOGRECORD*>( pBuffer );
            while ( ( pRecord->Length ) &&         // i.e. ensure always advance
                    ( idxOffset < bytesRead ) )
            {
               pRecord = reinterpret_cast<EVENTLOGRECORD*>(pBuffer + idxOffset);

               // Determine if this event is of the desired type or ID
               eventID = LOWORD( pRecord->EventID );
               if ( ( eventType == pRecord->EventType ) ||
                    ( PXSIsInUInt32Array( eventID, pEventIDs, numEventIDs) ) )
               {
                    Event.Set( pRecord );
                    pEventLogRecords->Append( Event );
                    numRecords = PXSAddSizeT( numRecords, 1 );
               }
               idxOffset = PXSAddUInt32( idxOffset, pRecord->Length );
            }

            // If reached maximum limit, set an information message
            if ( numRecords >= maximumRecords )
            {
                PXSLogAppInfo1( L"Reported maximum of %%1 event log records.",
                               Format.SizeT( numRecords ) );
                break;
            }
        }
    }
    catch ( const Exception& )
    {
        CloseEventLog( hEventLog );
        throw;
    }
    CloseEventLog( hEventLog );

    // Reset privilege if it was changed
    try
    {
        if ( privilegeChanged )
        {
            ProcessInfo.ChangePrivilege( SE_SECURITY_NAME, false );
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( L"Error resetting SE_SECURITY_NAME.", e, __FUNCTION__ );
    }
    pEventLogRecords->Rewind();
}

//===============================================================================================//
//  Description:
//      Get information about the specified exe file
//
//  Parameters:
//      FilePath      - the executable's file path
//      pFileName     - receives the file name
//      pPublisher    - receives the publisher name
//      pVersionString- receives the version string
//      pDescription  - receives the publisher's description
//
//  Returns:
//     void
//===============================================================================================//
void EventLogInformation::GetFileNameAndVersionInfo( const String& FilePath,
                                                     String* pFileName,
                                                     String* pPublisher,
                                                     String* pVersionString,
                                                     String* pDescription )
{
    File      FileObject;
    String    Drive, Dir, Ext;
    Directory DirObject;
    FileVersion FileVer;

    if ( ( pFileName      == nullptr ) ||
         ( pPublisher     == nullptr ) ||
         ( pVersionString == nullptr ) ||
         ( pDescription   == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }
    *pFileName      = PXS_STRING_EMPTY;
    *pPublisher     = PXS_STRING_EMPTY;
    *pVersionString = PXS_STRING_EMPTY;
    *pDescription   = PXS_STRING_EMPTY;

    // Test if have file name only or full path
    if ( FilePath.IndexOf( PXS_PATH_SEPARATOR, 0 ) == PXS_MINUS_ONE )
    {
        *pFileName = FilePath;
        return;
    }

    // Have the full path so can get version data
    try
    {
        if ( FileObject.Exists( FilePath ) &&
             FilePath.EndsWithStringI( L".exe" ) )
        {
            FileVer.GetVersion( FilePath, pPublisher, pVersionString, pDescription );
        }
        DirObject.SplitPath( FilePath, &Drive, &Dir, pFileName, &Ext );
        *pFileName += Ext;
    }
    catch ( const Exception& e )
    {
        PXSLogException( FilePath.c_str(), e, __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Get the message for the parameter ID by scanning the specified
//      message parameter files.
//
//  Parameters:
//      paramID          - the parameter identifier
//      ParameterFiles   - the files to look in
//      pParameterString - receives the parameter string, "" if not found
//
//  Returns:
//     void
//===============================================================================================//
void EventLogInformation::GetParameterString( DWORD paramID,
                                              const StringArray& ParameterFiles,
                                              String* pParameterString )

{
    File   FileObject;
    size_t i = 0;
    size_t numFiles = ParameterFiles.GetSize();
    String FilePath;
    Formatter   Format;
    StringArray EmptyArray;

    if ( pParameterString == nullptr )
    {
        throw ParameterException( L"pParameterString", __FUNCTION__ );
    }
    *pParameterString = PXS_STRING_EMPTY;

    while ( ( i < numFiles ) && pParameterString->IsEmpty() )
    {
        FilePath = ParameterFiles.Get( i );
        if ( FileObject.Exists( FilePath ) )
        {
            *pParameterString = Format.GetModuleString( FilePath, paramID, EmptyArray );
        }
        i++;
    };
}

//===============================================================================================//
//  Description:
//      Make the description associated with an event log entry
//
//  Parameters:
//      pszLogFile      - the name of the log file e.g. application
//      pEventLogRecord - the event log record
//      pDescription    - string object to receive the description
//
//  Remarks:
//      Problem:
//      MSDN docs show the EventID used as the message id with FormatMessage.
//      However, the EventID is the LOWORD of a message ID. If severity,
//      custom or facility bits are set then FormatMessage will fail when
//      using EventID alone. See "Reporting an Event". Looking in
//      some dll message tables there does not seem to be any consistent
//      behaviour as to whether or the higher bits are set. Additionally, there
//      appears to be no documentation on this will just use the EventID.
//
//      Problem 2:
//      XML data of the event record has the field "Provider Name"
//      which is presumably the documented "Source Name". This field also
//      has an attribute guid. This seems to point to the registry location
//      with the message, parameter and resource files, e.g.
//      HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\WINEVT\Publishers
//      \{54849625-5478-4994-a5ba-3e3b0328c30d}. Will stick to the
//      EVENTLOGRECORD structure documentation. In the event no description
//      can be made will return the EventID. For descriptions in the security
//      log see MSDN "Description of security events in Windows 7 and in
//      Windows Server 2008 R2"
//
//      For some messages in the event log, an insertion string may
//      contain a parameter requiring further substitution.
//      e.g. string in resource message table of netevent.dll (NT4 SP6)
//      id 2147490648 is:
//          'The %1 service failed to start due to the following error: %%2
//      In the event data structure:
//          string 1 = Remote Access Mac
//          string 2 = %%2
//      The form %%n in an event log record signifies that n is a code
//      to be obtained from the ParameterMessageFile registry value
//      MSDN: Article ID: Q125661
//
//  Returns:
//      void
//===============================================================================================//
void EventLogInformation::MakeEventLogDescription( LPCWSTR pszLogFile,
                                                   const EventLogRecord* pEventLogRecord,
                                                   String* pDescription )
{
    int       number = 0;
    File      FileObject;
    WORD      eventID = 0;
    DWORD     i = 0, maxSrings = 0, paramID = 0;
    size_t    index = 0, numEventFiles = 0, numInserts = 0, numIndexes = 0;
    String    OldString, SourceName, Insert, RegKey, KeyValue;
    String    FilePath, ParameterString, Temp;
    LPCWSTR   STRING_PP = L"%%";   // Parameter identifier in event log
    Registry  RegObject;
    Formatter Format;
    TArray< size_t > Indexes;
    StringArray StringInserts, ParameterFiles, EventFiles, EmptyArray;

    if ( pDescription == nullptr )
    {
        throw ParameterException( L"pDescription", __FUNCTION__ );
    }
    *pDescription = PXS_STRING_EMPTY;

    if ( ( pszLogFile == nullptr ) || ( pEventLogRecord == nullptr ) )
    {
        return;
    }

    // Get the insertion strings, 9 is the maximum on Vista+
    maxSrings = pEventLogRecord->GetNumberOfStrings();
    maxSrings = PXSMinUInt32( maxSrings, 9 );
    for ( i = 1; i <= maxSrings; i++ )
    {
        Insert = PXS_STRING_EMPTY;
        pEventLogRecord->GetString( i, &Insert );
        StringInserts.Add( Insert );
    }

    // Get the event and parameter message file names
    pEventLogRecord->GetSourceName( &SourceName );
    if ( SourceName.GetLength() )
    {
        // Construct the path to the key of the source name which is at the
        // end of the fixed size
        RegKey  = L"System\\CurrentControlSet\\Services\\EventLog\\";
        RegKey += pszLogFile;
        RegKey += L"\\";
        RegKey += SourceName;

        // Get the semi-colon delimited string of the event messages files
        // and parameter message files.
        RegObject.Connect( HKEY_LOCAL_MACHINE );
        RegObject.GetStringValue( RegKey.c_str(), L"EventMessageFile", &KeyValue);
        KeyValue.Trim();
        KeyValue.ToArray( ';', &EventFiles );
        RegObject.GetStringValue( RegKey.c_str(), L"ParameterMessageFile", &KeyValue );
        KeyValue.Trim();
        KeyValue.ToArray( ';', &ParameterFiles );
    }

    // Get the description, take the first one found in the message files
    i = 0;
    numEventFiles = EventFiles.GetSize();
    while ( ( i < numEventFiles ) && pDescription->IsEmpty() )
    {
        FilePath = EventFiles.Get( i );
        if ( FileObject.Exists( FilePath ) )
        {
            eventID = pEventLogRecord->GetEventID();
            *pDescription = Format.GetModuleString( FilePath, eventID, StringInserts );
        }
        else
        {
            PXSLogAppWarn1( L"Event Message File '%%1' not found.", FilePath);
        }
        i++;
    }

    // Substitute the insertion strings, will avoid FormatString when
    // formatting Event Log messages.
    numInserts = StringInserts.GetSize();
    for ( i = 0; i < numInserts; i++ )
    {
        OldString  = L"%";
        OldString += Format.UInt32( i + 1 );
        Insert = StringInserts.Get( i );
        pDescription->ReplaceI( OldString.c_str(), Insert.c_str() );
    }

    // Replace any %%n parameters
    Temp = *pDescription;
    Temp.IndexesOf( STRING_PP, false, &Indexes );
    numIndexes = Indexes.GetSize();
    for ( i = 0; i < numIndexes; i++ )
    {
        index  = Indexes.Get( i );
        index  = PXSAddSizeT( index, 2 );         // Skip over the %%
        number = _wtoi( Temp.c_str() + index );
        if ( number >= 0 )
        {
            paramID = PXSCastInt32ToUInt32( number );
            GetParameterString( paramID, ParameterFiles, &ParameterString );
            if ( ParameterString.GetLength() > 0 )
            {
                // Replace with the parameter string.
                OldString  = L"%%";
                OldString += Format.UInt32( paramID );
                pDescription->ReplaceI( OldString.c_str(),
                                        ParameterString.c_str() );
            }
        }
    }
    pDescription->Trim();

    // If no description, show the event id, may be zero or perhaps
    // no message file, note displaying the the low word part
    if ( pDescription->IsEmpty() )
    {
        eventID = pEventLogRecord->GetEventID();
        *pDescription  = L"Event ID = ";
        *pDescription += Format.UInt16( eventID );
        pDescription->Trim();
    }
}
