///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Event Log Record Class Implementation
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
#include "WinAudit/Header Files/EventLogRecord.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StringT.h"

// 5. This Project

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
EventLogRecord::EventLogRecord()
               :m_pEventLogRecord( nullptr )
{
}

// Copy constructor
EventLogRecord::EventLogRecord( const EventLogRecord& oEventLogRecord )
               :m_pEventLogRecord( nullptr )
{
    *this = oEventLogRecord;
}

// Destructor
EventLogRecord::~EventLogRecord()
{
    if ( m_pEventLogRecord )
    {
        delete [] m_pEventLogRecord;    // Allocated as byte[] array
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator
EventLogRecord& EventLogRecord::operator= ( const EventLogRecord& oEventLogRecord )
{
    if ( this == &oEventLogRecord ) return *this;

    Set( oEventLogRecord.m_pEventLogRecord );

    return *this;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the event ID associated with this event record
//
//  Parameters:
//      None
//
//  Remarks:
//      The EventID member of EVENTLOGRECORD is a DWORD however, will return
//      the low WORD as this is what is most often used.
//
//  Returns:
//      time_t
//===============================================================================================//
WORD EventLogRecord::GetEventID() const
{
    if ( m_pEventLogRecord == nullptr )
    {
        throw FunctionException( L"m_pEventLogRecord", __FUNCTION__ );
    }

    return PXSCastUInt32ToUInt16( 0xFFFF & m_pEventLogRecord->EventID );
}

//===============================================================================================//
//  Description:
//      Get the number of strings in this event record
//
//  Parameters:
//      None
//
//  Returns:
//      WORD
//===============================================================================================//
WORD EventLogRecord::GetNumberOfStrings() const
{
    if ( m_pEventLogRecord == nullptr )
    {
        throw FunctionException( L"m_pEventLogRecord", __FUNCTION__ );
    }

    return m_pEventLogRecord->NumStrings;
}

//===============================================================================================//
//  Description:
//      Get event log record's source name
//
//  Parameters:
//      pSourceName - receives the source name
//
//  Returns:
//     void
//===============================================================================================//
void EventLogRecord::GetSourceName( String* pSourceName ) const
{
    size_t  i = 0;
    wchar_t ch;
    BYTE*   pByte;
    LPCWSTR psz;

    if ( m_pEventLogRecord == nullptr )
    {
        throw FunctionException( L"m_pEventLogRecord", __FUNCTION__ );
    }

    if ( pSourceName == nullptr )
    {
        throw ParameterException( L"pSourceName", __FUNCTION__ );
    }
    *pSourceName = PXS_STRING_EMPTY;

    // The source name is immediately after the fixed part EVENTLOGRECORD
    // structure
    if ( m_pEventLogRecord->Length <= sizeof ( EVENTLOGRECORD ) )
    {
        return;     // Nothing after the fixed part
    }

    // Presumably the SourceName[] data is terminated but the documentation
    // does not say this so will test for the first known offset which is
    // UserSidOffset so as not to run past that.
    pSourceName->Allocate( MAX_PATH );
    pByte = reinterpret_cast<BYTE*>( m_pEventLogRecord );
    psz   = reinterpret_cast<LPCWSTR>( pByte + sizeof ( EVENTLOGRECORD ) );
    ch    = *psz;
    while ( ( ch != PXS_CHAR_NULL ) &&
            ( i  <  m_pEventLogRecord->UserSidOffset ) &&
            ( i  <  m_pEventLogRecord->Length ) )
    {
        *pSourceName += ch;
        i++;
        psz++;
        ch = *psz;
    }
}

//===============================================================================================//
//  Description:
//      Extract string data from the specified event record
//
//  Parameters:
//      stringNumber  - the string number
//      pString       - receives the string
//
//  Remarks:
//      The maximum number of string in a record is 6 pre-Vista and 9 onVista+
//
//  Returns:
//      void, pString receives "" if the string is not found
//===============================================================================================//
void EventLogRecord::GetString( DWORD stringNumber, String* pString ) const

{
    DWORD   start    = 0, numChars = 0, index = 0, count = 0;
    BYTE*   pByte    = nullptr;
    LPCWSTR pStrings = nullptr;

    if ( pString == nullptr )
    {
        throw ParameterException( L"pString", __FUNCTION__ );
    }
    *pString = PXS_STRING_EMPTY;

    if ( m_pEventLogRecord == nullptr )
    {
        return;     // Nothing to do or error
    }

    if ( stringNumber > m_pEventLogRecord->NumStrings )
    {
        return;     // Nothing to do
    }

    // Step through the string data. Strings[] precede the Data[] block
    if ( m_pEventLogRecord->StringOffset >= m_pEventLogRecord->DataOffset )
    {
        return;     // No data
    }
    numChars = (m_pEventLogRecord->DataOffset - m_pEventLogRecord->StringOffset) / sizeof(wchar_t);
    pByte    = reinterpret_cast<BYTE*>( m_pEventLogRecord );
    pStrings = reinterpret_cast<LPCWSTR>(pByte+m_pEventLogRecord->StringOffset);
    while ( ( index < numChars ) && ( count < stringNumber ) )
    {
        if ( pStrings[ index ] == PXS_CHAR_NULL )
        {
            count++;
            if ( stringNumber == count )
            {
                *pString = pStrings + start;
            }
            index++;
            start = index;          // Start of the next string
        }
        else
        {
            index++;
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the time the log record was generated
//
//  Parameters:
//      None
//
//  Returns:
//      time_t
//===============================================================================================//
time_t EventLogRecord::GetTimeGenerated() const
{
    if ( m_pEventLogRecord == nullptr )
    {
        throw FunctionException( L"m_pEventLogRecord", __FUNCTION__ );
    }

    return PXSCastUInt32ToTimeT( m_pEventLogRecord->TimeGenerated );
}

//===============================================================================================//
//  Description:
//      Set this event log record based on the input EVENTLOGRECORD structure
//
//  Parameters:
//      None
//
//  Returns:
//      time_t
//===============================================================================================//
void EventLogRecord::Set( const EVENTLOGRECORD* pEventLogRecord )
{
    BYTE* pNew;

    // Special case of NULL
    if ( pEventLogRecord == nullptr )
    {
        delete [] m_pEventLogRecord;    // Allocated as byte[] array
        m_pEventLogRecord = nullptr;
        return;
    }

    if ( pEventLogRecord->Length < sizeof ( EVENTLOGRECORD ) )
    {
        // Should not get here unless EVENTLOGRECORD has bad data
        throw ParameterException( L"pEventLogRecord->Length", __FUNCTION__ );
    }

    pNew = new BYTE[ pEventLogRecord->Length ];
    if ( pNew == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    memcpy( pNew, pEventLogRecord, pEventLogRecord->Length );

    // Replace
    delete [] m_pEventLogRecord;
    m_pEventLogRecord = reinterpret_cast<EVENTLOGRECORD*>( pNew );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
