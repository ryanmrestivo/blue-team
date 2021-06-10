///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Logger Class Header
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

#ifndef PXSBASE_LOGGER_H_
#define PXSBASE_LOGGER_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

// Logging can be to a file, a string array or both.

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/PxsBase.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/StringT.h"
#include "PxsBase/Header Files/StringArray.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class Logger
{
    public:
        // Constructor
        Logger();

        // Destructor
        ~Logger();

        // Methods
        void    AppendComment( LPCWSTR pszComment );
        void    Flush();
        void    GetLogMessages( StringArray* pLogMessages ) const;
const String&   GetPath() const;
        bool    IsStarted() const;
        void    PurgeLogMessages();
        DWORD   Start( bool wantExtraInfo );
        DWORD   Start( const String& Path, bool createNew );
        void    Stop();
        void    SetWantComputerName( bool wantComputerName );
        void    WriteEntry( DWORD severity,
                            DWORD errorType,
                            DWORD errorCode,
                            bool translateError,
                            LPCWSTR pszMessage, const String& Insert1, const String& Insert2 );
    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        Logger( const Logger& oLogger );

        // Assignment operator - not allowed
        Logger& operator= ( const Logger& oLogger );

        // Methods
        void    CreateEntry( DWORD severity,
                             DWORD derrorType,
                             DWORD errorCode,
                             bool translateError,
                             LPCWSTR pszMessage,
                             const String& Insert1, const String& Insert2, String* pEntry );
        void    SeverityCodeToString( DWORD severity, String* pSeverityString );
        void    WriteText( const String& Text );
        void    WriteHeader();
        void    WriteSystemDiagnosticData( bool wantExtraInfo );

        // Data members
        bool    m_bLogging;
        bool    m_bWantComputerName;
        bool    m_bWritingEntry;       // Indicate an entry is being written
        File    m_File;
        String  m_FilePath;
        StringArray m_LogMessages;     // Array to hold the messages
};

#endif  // PXSBASE_LOGGER_H_
