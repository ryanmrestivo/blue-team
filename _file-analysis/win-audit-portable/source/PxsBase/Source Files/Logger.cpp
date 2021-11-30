///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Log to a File Class Implementation
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

// Logging public methods do not throw exceptions. Errors are simply ignored
// as there is is no one to tell. Also don't want recursion.

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/Logger.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/SystemInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

//
// Default constructor, an empty logging object
//
Logger::Logger()
       :m_bLogging( false ),
        m_bWantComputerName( false ),
        m_bWritingEntry( false ),
        m_File(),
        m_FilePath(),
        m_LogMessages()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
Logger::~Logger()
{
    try
    {
        Stop();
    }
    catch ( const Exception& )
    { }     // Ignore
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
//      Add a comment line to the log file
//
//  Parameters:
//      pszComment - comment to add
//
//  Returns:
//      void
//===============================================================================================//
void Logger::AppendComment( LPCWSTR pszComment )
{
    String Comment;

    try
    {
        if ( IsStarted() )
        {
            Comment  = L"!";
            Comment += pszComment;
            WriteText( Comment );
        }
    }
    catch ( const Exception& )
    { }     // Ignore
}

//===============================================================================================//
//  Description:
//      Flush the logger output stream. Only applies to the log file.
//      Has no effect if logging to an array.
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Logger::Flush()
{
    try
    {
        m_File.Flush();
    }
    catch ( const Exception& )
    { }     // Ignore
}

//===============================================================================================//
//  Description:
//      Determine if the logger has been started.
//
//  Parameters:
//      None
//
//  Returns:
//      true if started, else false
//===============================================================================================//
bool Logger::IsStarted() const
{
    return m_bLogging;
}

//===============================================================================================//
//  Description:
//      Get the log messages stored in the string array
//
//  Parameters:
//      pLogMessages - receives the log messages
//
//  Returns:
//      void
//===============================================================================================//
void Logger::GetLogMessages( StringArray* pLogMessages ) const
{
    try
    {
        if ( pLogMessages )
        {
            *pLogMessages = m_LogMessages;
        }
    }
    catch ( const Exception& )
    { }     // Ignore
}

//===============================================================================================//
//  Description:
//      Get the log file path. Only applies if logging to a file.
//
//  Parameters:
//      None
//
//  Returns:
//      Constant reference to a string object
//===============================================================================================//
const String& Logger::GetPath() const
{
    return m_FilePath;
}

//===============================================================================================//
//  Description:
//      Purge the log messages stored in the string array
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Logger::PurgeLogMessages()
{
    m_LogMessages.RemoveAll();
}

//===============================================================================================//
//  Description:
//      Start logging to a string array.
//
//  Parameters:
//      wantExtraInfo - indicates want extra dignostic messages
//
//  Returns:
//      System error code
//===============================================================================================//
DWORD Logger::Start( bool wantExtraInfo )
{
    String Line;

    if ( m_bLogging )
    {
        return ERROR_SUCCESS;     // Nothing to do
    }
    m_bLogging = true;

    // If already have messages, add a few blank lines
    if ( m_LogMessages.GetSize()  )
    {
        Line += L"!\r\n";
        Line += L"!\r\n";
        Line += L"!Logging started.";
        WriteText( Line );
    }
    else
    {
        WriteHeader();
    }
    WriteSystemDiagnosticData( wantExtraInfo );

    return ERROR_SUCCESS;
}

//===============================================================================================//
//  Description:
//      Start logging to a file
//
//  Parameters:
//      Path      - path to log file
//      createNew - flag for create new log file, else append
//
//  Remarks:
//      Does not throw exceptions, returns an error code.
//      Create new and header flags are only used if creating a new log file.
//
//  Returns:
//      DWORD error code
//===============================================================================================//
DWORD Logger::Start( const String& Path, bool createNew )
{
    bool   exists    = false;
    DWORD  errorCode = 0;
    String NewLine;

    try
    {
        // If already started then return.
        if ( m_File.IsOpen() )
        {
            return ERROR_SUCCESS;
        }

        // Must have set the log path
        m_FilePath = Path;
        m_FilePath.Trim();
        if ( m_FilePath.IsEmpty() )
        {
            return ERROR_BAD_PATHNAME;
        }
        m_bLogging = true;

        // Create a new log file if none exists or want a new one
        exists = m_File.Exists( m_FilePath );
        if ( createNew || ( exists == false ) )
        {
            // Delete any existing file
            if ( exists )
            {
                m_File.Delete( m_FilePath );
            }

            // Create a new one with share access
            m_File.CreateNew( m_FilePath, FILE_SHARE_READ, true );
            WriteHeader();
        }
        else
        {
            // Open existing text file
            m_File.Open( m_FilePath, GENERIC_READ | GENERIC_WRITE, FILE_SHARE_READ, 1, true );
            m_File.SeekToEnd();
            NewLine = PXS_STRING_CRLF;
            WriteText( NewLine );
        }
        WriteSystemDiagnosticData( true );
    }
    catch ( const Exception& e )
    {
        errorCode  = e.GetErrorCode();
    }

    return errorCode;
}

//===============================================================================================//
//  Description:
//      Stop logging
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Logger::Stop()
{
    try
    {
        String Line;

        Line = L"!\r\n!Logging stopped.\r\n!";
        WriteText( Line );
        m_File.Close();
        m_bLogging = false;
    }
    catch ( const Exception& )
    { }     // Ignore
}

//===============================================================================================//
//  Description:
//      Set if want to include the column of computer name
//
//  Parameters:
//        wantComputerName - if true will include the computer's name
//
//  Returns:
//      void
//===============================================================================================//
void Logger::SetWantComputerName( bool wantComputerName )
{
    m_bWantComputerName = wantComputerName;
}

//===============================================================================================//
//  Description:
//      Write a log entry in the log file
//
//  Parameters:
//      severity      - defined constant of severity e.g. info, warning etc.
//      errorType     - type of error, e.g. system, network etc.
//      errorCode     - a numerical code of the error
//      translateError- flag to indicate want to translate the error code
//      pszMessage    - string message with optional substitution parameters
//      Insert1       - string insert parameter, any %%1 are replaced
//      Insert2       - string insert parameter, any %%2 are replaced
//
//  Returns:
//      void
//===============================================================================================//
void Logger::WriteEntry( DWORD severity,
                         DWORD errorType,
                         DWORD errorCode,
                         bool translateError,
                         LPCWSTR pszMessage, const String& Insert1, const String& Insert2 )
{
    String Entry;

    // Can only write one entry at a time
    if ( m_bWritingEntry )
    {
        return;
    }

    try
    {
        // Must be logging
        if ( IsStarted() == false )
        {
            return;
        }

        // Set the class scope flag to say writing to the log file, must reset
        // this flag at the end of this method so that other
        m_bWritingEntry = true;
        CreateEntry( severity,
                     errorType, errorCode, translateError, pszMessage, Insert1, Insert2, &Entry );
        WriteText( Entry );

        // For severity error level, flush to ensure it is written
        if ( severity <= PXS_SEVERITY_ERROR )
        {
            Flush();
        }
    }
    catch ( const Exception& )
    { }     // Ignore

    m_bWritingEntry = false;    // Reset flag
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Create a formatted single line log entry
//
//  Parameters:
//      severity   - defined constant of severity e.g. info, warning etc.
//      errorType  - type of error, e.g. system, network etc.
//      errorCode  - a numerical code of the error
//      pszMessage - string message with optional substitution parameters
//      Insert1    - string insert parameter, any %%1 are replaced
//      Insert2    - string insert parameter, any %%2 are replaced
//      pEntry     - string object to receive the entry
//
//  Returns:
//      void
//===============================================================================================//
void Logger::CreateEntry( DWORD severity,
                          DWORD errorType,
                          DWORD errorCode,
                          bool translateError,
                          LPCWSTR pszMessage,
                          const String& Insert1, const String& Insert2, String* pEntry )
{
    String     Temp, Translation, ErrorName;
    Formatter  Format;
    Exception  eTranslate;
    SystemInformation SystemInfo;

    if ( pEntry == nullptr )
    {
        return;
    }

    try
    {
        *pEntry = PXS_STRING_EMPTY;
        pEntry->Allocate( 512 );

        // Start with time stamp
        *pEntry  = Format.NowTohhmmssttt();
        *pEntry += PXS_CHAR_TAB;

        // Computer Name
        if ( m_bWantComputerName )
        {
            SystemInfo.GetComputerNetBiosName( &Temp );
            Temp.FixedWidth( 11, PXS_CHAR_SPACE );
            *pEntry += Temp;
            *pEntry += PXS_CHAR_TAB;
        }

        // Severity
        Temp = PXS_STRING_EMPTY;
        SeverityCodeToString( severity, &Temp );
        Temp.FixedWidth( 11, PXS_CHAR_SPACE );
        *pEntry += Temp;
        *pEntry += PXS_CHAR_TAB;

        // Error Type
        eTranslate.TranslateType( errorType, &ErrorName );
        Temp = ErrorName;
        Temp.FixedWidth( 11, PXS_CHAR_SPACE );
        *pEntry += Temp;
        *pEntry += PXS_CHAR_TAB;

        // Error Code, if its a large number show as hex
        Temp = PXS_STRING_EMPTY;
        if ( errorCode > 0xFFFF )
        {
            Temp = Format.UInt32Hex( errorCode, true );
        }
        else
        {
            Temp = Format.UInt32( errorCode );
        }
        Temp.FixedWidth( 11, PXS_CHAR_SPACE );
        *pEntry += Temp;
        *pEntry += PXS_CHAR_TAB;

        // Add the Message, format the optional string parameters
        if ( pszMessage )
        {
            *pEntry += Format.String2( pszMessage, Insert1, Insert2 );
        }

        // Translate the error code, zero always means success
        if ( errorCode && translateError )
        {
            eTranslate.Translate( errorType, errorCode, &Translation );

            // Add to line
            *pEntry += PXS_CHAR_SPACE;
            *pEntry += Translation;
        }

        // Put on one line
        pEntry->ReplaceChar( PXS_CHAR_LF, PXS_CHAR_SPACE );
        pEntry->ReplaceChar( PXS_CHAR_CR, PXS_CHAR_SPACE );
        pEntry->Trim();
    }
    catch ( const Exception& )
    { }     // Ignore
}

//===============================================================================================//
//  Description:
//      Translate the specified severity code to a string
//
//  Parameters:
//      severity        - the severity code
//      pSeverityString - receives the English translation
//
//  Returns:
//      void
//===============================================================================================//
void Logger::SeverityCodeToString( DWORD severity, String* pSeverityString )
{
    Formatter Format;

    if ( pSeverityString == nullptr )
    {
        return;
    }

    try
    {
        *pSeverityString = PXS_STRING_EMPTY;

        if ( severity == PXS_SEVERITY_ERROR )
        {
            *pSeverityString = L"Error";
        }
        else if ( severity == PXS_SEVERITY_WARNING )
        {
            *pSeverityString = L"Warning";
        }
        else if ( severity == PXS_SEVERITY_INFORMATION )
        {
            *pSeverityString = L"Information";
        }
        else
        {
            *pSeverityString = Format.UInt32( severity );
        }
    }
    catch ( const Exception& )
    { }     // Ignore
}

//===============================================================================================//
//  Description:
//      Write the log file's header
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void Logger::WriteHeader()
{
    String ApplicationName, Header;

    try
    {
        Header.Allocate( 512 );
        Header  = L"!========================================================"
                  L"========================================================="
                  L"=====================================================\r\n";
        Header += L"!H:MM:SS.ttt\t";
        if ( m_bWantComputerName )
        {
            Header += L"COMPUTER   \t";
        }
        Header += L"SEVERITY   \tTYPE       \tNUMBER     \tDESCRIPTION\r\n"
                  L"!========================================================"
                  L"========================================================="
                  L"=====================================================";
        WriteText( Header );
    }
    catch ( const Exception& )
    { }     // Ignore
}

//===============================================================================================//
//  Description:
//      Write diagnostic data to the log file
//
//  Parameters:
//      wantExtraInfo - indicates want extrq diagnostic messages
//
//  Returns:
//      void
//===============================================================================================//
void Logger::WriteSystemDiagnosticData( bool wantExtraInfo )
{
    size_t    i = 0, numElements = 0;
    String    ComputerColumn, Line;
    Formatter Format;
    StringArray DiagnosticData;
    SystemInformation SystemInfo;

    try
    {
        ComputerColumn = PXS_STRING_EMPTY;
        if ( m_bWantComputerName )
        {
            SystemInfo.GetComputerNetBiosName( &ComputerColumn );
            ComputerColumn += PXS_CHAR_TAB;
        }

        if ( g_pApplication )
        {
            g_pApplication->GetDiagnosticData( wantExtraInfo, &DiagnosticData );
        }
        Line.Allocate( 512 );
        numElements = DiagnosticData.GetSize();
        for ( i = 0; i < numElements; i++ )
        {
            Line  = Format.NowTohhmmssttt();
            Line += PXS_CHAR_TAB;
            Line += ComputerColumn;
            Line += L"Information";
            Line += PXS_CHAR_TAB;
            Line += L"Application";
            Line += PXS_CHAR_TAB;
            Line += PXS_STRING_ZERO;
            Line += PXS_CHAR_TAB;
            Line += PXS_CHAR_TAB;
            Line += DiagnosticData.Get( i );
            WriteText( Line );
        }
    }
    catch ( const Exception& )
    { }     // Ignore
}

//===============================================================================================//
//  Description:
//      Write the specified line to the log file and/or string array
//
//  Parameters:
//      Text - the line to write
//
//  Returns:
//      void
//===============================================================================================//
void Logger::WriteText( const String& Text )
{
    String NewLine;

    if ( ( m_bLogging == false ) || ( Text.IsEmpty() ) )
    {
        return;
    }

    try
    {
        if ( m_File.IsOpen() )
        {
            m_File.WriteChars( Text );
            NewLine = PXS_STRING_CRLF;
            m_File.WriteChars( NewLine );
        }
        m_LogMessages.Add( Text );
    }
    catch ( const Exception& )
    { }     // Ignore
}
