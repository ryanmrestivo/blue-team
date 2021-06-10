///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Event Log Records Class Header
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

#ifndef WINAUDIT_EVENT_LOG_RECORD_H_
#define WINAUDIT_EVENT_LOG_RECORD_H_

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

// 5. This Project

// 6. Forwards
class String;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class EventLogRecord
{
    public:
        // Default constructor
        EventLogRecord();

        // Destructor
        ~EventLogRecord();

        // Copy constructor
        EventLogRecord( const EventLogRecord& oEventLogRecord );

        // Assignment operator
        EventLogRecord& operator= ( const EventLogRecord& oEventLogRecord );

        // Methods
        WORD    GetEventID() const;
        WORD    GetNumberOfStrings() const;
        void    GetString( DWORD stringNumber, String* pString ) const;
        void    GetSourceName( String* pSourceName ) const;
        time_t  GetTimeGenerated() const;
        void    Set( const EVENTLOGRECORD* pEventLogRecord );

        // Members

    protected:
        // Methods

        // Members

    private:
        // Methods

        // Members
        EVENTLOGRECORD* m_pEventLogRecord;
};

#endif  // WINAUDIT_EVENT_LOG_RECORD_H_
