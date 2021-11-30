///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Event Log Information Class Header
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

#ifndef WINAUDIT_EVENT_LOG_INFORMATION_H_
#define WINAUDIT_EVENT_LOG_INFORMATION_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/WinAudit.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/TList.h"

// 5. This Project
#include "WinAudit/Header Files/EventLogRecord.h"

// 6. Forwards
class AuditRecord;
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class EventLogInformation
{
    public:
        // Default constructor
        EventLogInformation();

        // Destructor
        ~EventLogInformation();

        // Methods
        void    GetAuditFailureRecords( size_t maximumEntries, TArray< AuditRecord >* pRecords );
        void    GetErrorRecords( size_t maximumEntries, TArray< AuditRecord >* pRecords );
        void    GetSoftwareMeteringRecords( TArray< AuditRecord >* pRecords );
        void    GetSourceNames( StringArray* pSourceNames );
        void    GetUptimeStatsRecord( AuditRecord* pRecord );
        void    GetUserLogonRecords( TArray< AuditRecord >* pRecords );

    protected:
        // Methods

        // Data members

    private:
        typedef struct _TYPE_METERING_DATA
        {
            time_t  firstStartTime;
            time_t  lastStartTime;
            DWORD   numberStarts;
            wchar_t szFileName[ MAX_PATH + 1 ];
        } TYPE_METERING_DATA;

        typedef struct _TYPE_USER_DATA
        {
            time_t  firstLogonTime;
            time_t  lastLogonTime;
            DWORD   consoleLogons;    // Event 582, type 2
            DWORD   remoteLogons;     // Event 582, type 10, require XP
            DWORD   otherLogons;      // Machine account, batch, network, etc.
            wchar_t szUserName[ UNLEN + 1 ];
        } TYPE_USER_DATA;

        // Copy constructor - not allowed
        EventLogInformation( const EventLogInformation& oEventLogInformation );

        // Assignment operator - not allowed
        EventLogInformation& operator= ( const EventLogInformation& oEventLogInformation );

        // Methods
        bool FindMeteringData( const String& FileName,
                               TList< TYPE_METERING_DATA >* pMeteringData,
                               TYPE_METERING_DATA* pElement );

        bool FindUserData( const String& DomainUserName,
                           TList< TYPE_USER_DATA >* PUserData, TYPE_USER_DATA* pElement );
        void GetSourceNameAuditRecords( DWORD auditCategoryID,
                                        DWORD timeGeneratedItemID,
                                        DWORD logFileItemID,
                                        DWORD sourceItemID,
                                        DWORD decriptionItemID,
                                        LPCWSTR pszLogFile,
                                        WORD  eventType,
                                        size_t maximumEntries, TArray<AuditRecord>* pRecords );
        void GetEventLogRecords(LPCWSTR pszName,
                                bool forwardsRead,
                                const DWORD* EventIDs,
                                size_t numEventIDs,
                                WORD  eventType,
                                size_t maximumRecords, TList<EventLogRecord>* pEventLogRecords );
        void GetFileNameAndVersionInfo( const String& FilePath,
                                        String* pFileName,
                                        String* pPublisher,
                                        String* pVersionString, String* pDescription );
        void GetParameterString( DWORD paramID,
                                 const StringArray& ParameterFiles, String* pParameterString );
        void MakeEventLogDescription( LPCWSTR pszLogFile,
                                      const EventLogRecord* pEventLogRecord,
                                      String* pDescription );
        // Data members
};

#endif  // WINAUDIT_EVENT_LOG_INFORMATION_H_
