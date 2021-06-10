///////////////////////////////////////////////////////////////////////////////////////////////////
//
// ODBC Database and Wrapper Class Implementation
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

// Some databases only allow one active statement on a connection so will
// allocate a class scope handle for that purpose and re-use it when required.
//
// Type conversions from SQLCHAR (unsigned char) to char have been removed.
// This file should be compiled with UNICODE/_UNICODE otherwise will receive
// a number or warnings about sign conversions.
//
///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/OdbcDatabase.h"

// 2. C System Files
#include <sqlext.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/SystemInformation.h"

// 5. This Project
#include "WinAudit/Header Files/OdbcRecordSet.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
OdbcDatabase::OdbcDatabase()
             :m_EnvironmentHandle( nullptr ),
              m_ConnectionHandle( nullptr ),
              m_hTransactionStatement( nullptr ),
              m_ODBC(),
              m_bConnected( false ),
              m_uQueryTimeOut( 60 ),
              m_DbmsName(),
              m_DbmsVersion(),
              m_DmVersion(),
              m_DriverName(),
              m_DriverVersion(),
              m_OdbcVersion(),
              m_DataSourceName(),
              m_ServerName(),
              m_UserName()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
OdbcDatabase::~OdbcDatabase()
{
    try
    {
        Disconnect();
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
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
//      Begin a transaction
//
//  Parameters:
//      None
//
//  Remarks:
//      Nested transaction are not allowed.
//      Creates a class level statement object.
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::BeginTrans()
{
    // Must be connected
    if ( IsConnected() == false )
    {
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_NOT_CONNECTED, PXS_STRING_EMPTY, __FUNCTION__ );
    }

    // Test for an existing transaction
    if ( StartedTrans() )
    {
        return;     // A transaction is in progress
    }

    m_ODBC.SetConnectAttr( m_ConnectionHandle,
                           SQL_ATTR_AUTOCOMMIT,
                           nullptr,             // SQL_AUTOCOMMIT_OFF = 0
                           SQL_IS_UINTEGER );

    FreeTransactionStatement();     // Reuse the statement
}

//===============================================================================================//
//  Description:
//      Commit a transaction
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::CommitTrans()
{
    Formatter Format;

    // Verify connected to database
    if ( IsConnected() == false )
    {
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_NOT_CONNECTED, PXS_STRING_EMPTY, __FUNCTION__ );
    }

    if ( StartedTrans() == false )
    {
        return;     // Nothing to do
    }

    FreeTransactionStatement();     // Reuse the statement
    m_ODBC.EndTran( SQL_HANDLE_DBC, m_ConnectionHandle, SQL_COMMIT );
    m_ODBC.SetConnectAttr( m_ConnectionHandle,
                           SQL_ATTR_AUTOCOMMIT,
                           reinterpret_cast<SQLPOINTER>( SQL_AUTOCOMMIT_ON ),
                           SQL_IS_UINTEGER );
}

//===============================================================================================//
//  Description:
//      Connect to a database using the specified connection string
//
//  Parameters:
//      ConnectionString   - the full connection string;
//      connectTimeoutSecs - the login and connection timeout in seconds
//      queryTimeoutSecs   - the query timeout in seconds
//
//  Remarks:
//      The "OutConnectionString" is not returned to the caller
//      Workaround: On Windows 10, after an update of Office 365,
//      there is an ODBC error in the driver manager when using
//      SQL_DRIVER_COMPLETE. This happens even if there the DB is not password
//      protected so no additional connection information should be required.
//      It does work if use SQL_DRIVER_NOPROMPT so will force that. This means
//      cannot connect to a password protected Access DB on Win10. However, can
//      still do so at the command line as can specify the password.
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::Connect( const String& ConnectionString,
                            DWORD connectTimeoutSecs,
                            DWORD queryTimeoutSecs, HWND hWnd )
{
    bool isAccessOnWin10 = false;
    String       Temp, DatabaseName, OutConnectionString;
    SQLWCHAR     InfoValue[ 512 ] = { 0 };      // N.B. must an even number
    Formatter    Format;
    SQLSMALLINT  StringLength = 0;
    SQLUSMALLINT DriverCompletion = SQL_DRIVER_NOPROMPT;
    SystemInformation SysInfo;

    if ( ConnectionString.IsEmpty() )
    {
       throw ParameterException( L"ConnectionString", __FUNCTION__ );
    }

    // If connected, then disconnect to roll back and releases all resources
    if ( IsConnected() )
    {
        Disconnect();
    }

    m_ODBC.AllocHandle( SQL_HANDLE_ENV, nullptr, &m_EnvironmentHandle );
    m_ODBC.SetEnvAttr( m_EnvironmentHandle,
                       SQL_ATTR_ODBC_VERSION,
                       reinterpret_cast<SQLPOINTER>( SQL_OV_ODBC3 ),
                       SQL_IS_INTEGER );
    m_ODBC.AllocHandle( SQL_HANDLE_DBC,
                        m_EnvironmentHandle, &m_ConnectionHandle );

    // Will use the same value for login and connection timeouts
    try
    {
        DWORD_PTR dwordPtr = static_cast<DWORD_PTR>( connectTimeoutSecs );
        m_ODBC.SetConnectAttr(
                        m_ConnectionHandle,
                        SQL_ATTR_LOGIN_TIMEOUT,
                        reinterpret_cast<SQLPOINTER>( dwordPtr ),
                        SQL_IS_UINTEGER );

        m_ODBC.SetConnectAttr(
                        m_ConnectionHandle,
                        SQL_ATTR_CONNECTION_TIMEOUT,
                        reinterpret_cast<SQLPOINTER>( dwordPtr ),
                        SQL_IS_UINTEGER );
    }
    catch ( const Exception& e )
    {
        // Log and continue
        PXSLogException( L"Failed to set timeouts", e, __FUNCTION__ );
    }

    // Determine if connecting to Access on Windows 10
    if ( ( SysInfo.GetMajorVersion() >= 10 ) &&
         ( ConnectionString.IndexOfI( L"Microsoft Access Driver" ) != PXS_MINUS_ONE ) )
    {
        isAccessOnWin10 = true;
    }

    // Prompt if have a window as long as not Access on Windows 10
    if ( ( hWnd != nullptr ) && ( isAccessOnWin10 == false ) )
    {
        DriverCompletion = SQL_DRIVER_COMPLETE;
    }
    m_ODBC.DriverConnect( m_ConnectionHandle,
                          hWnd, ConnectionString, &OutConnectionString, DriverCompletion );
    m_bConnected = true;
    PXSLogAppInfo1( L"Completion connection string (-PWD): '%%1'.",
                    OutConnectionString );

    // Area under dialog sometimes needs redrawing
    if ( hWnd )
    {
        RedrawWindow( hWnd,
                      nullptr,
                      nullptr, RDW_ERASE | RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN );
    }

    // Allocate a statement handle for use in transactions
    m_ODBC.AllocHandle( SQL_HANDLE_STMT, m_ConnectionHandle, &m_hTransactionStatement );
    m_uQueryTimeOut = queryTimeoutSecs;
    SetQueryTimeoutAttr( m_hTransactionStatement );

    // DBMS name
    memset( InfoValue, 0, sizeof ( InfoValue ) );
    StringLength = 0;
    m_ODBC.GetInfo( m_ConnectionHandle,
                    SQL_DBMS_NAME, InfoValue, ARRAYSIZE( InfoValue ), &StringLength );
    InfoValue[ ARRAYSIZE( InfoValue ) - 1 ] = PXS_CHAR_NULL;
    m_DbmsName = InfoValue;
    m_DbmsName.Trim();

    // DBMS Version
    memset( InfoValue, 0, sizeof ( InfoValue ) );
    StringLength = 0;
    m_ODBC.GetInfo( m_ConnectionHandle,
                    SQL_DBMS_VER, InfoValue, ARRAYSIZE( InfoValue ), &StringLength );
    InfoValue[ ARRAYSIZE( InfoValue ) - 1 ] = PXS_CHAR_NULL;
    m_DbmsVersion = InfoValue;
    m_DbmsVersion.Trim();

    // Driver Manager Version
    memset( InfoValue, 0, sizeof ( InfoValue ) );
    StringLength = 0;
    m_ODBC.GetInfo( m_ConnectionHandle,
                    SQL_DM_VER, InfoValue, ARRAYSIZE( InfoValue ), &StringLength );
    InfoValue[ ARRAYSIZE( InfoValue ) - 1 ] = PXS_CHAR_NULL;
    m_DmVersion = InfoValue;
    m_DmVersion.Trim();

    // Driver Name
    memset( InfoValue, 0, sizeof ( InfoValue ) );
    StringLength = 0;
    m_ODBC.GetInfo( m_ConnectionHandle,
                    SQL_DRIVER_NAME, InfoValue, ARRAYSIZE( InfoValue ), &StringLength );
    InfoValue[ ARRAYSIZE( InfoValue ) - 1 ] = PXS_CHAR_NULL;
    m_DriverName = InfoValue;
    m_DriverName.Trim();

    // Driver Version
    memset( InfoValue, 0, sizeof ( InfoValue ) );
    StringLength = 0;
    m_ODBC.GetInfo( m_ConnectionHandle,
                    SQL_DRIVER_VER, InfoValue, ARRAYSIZE( InfoValue ), &StringLength );
    InfoValue[ ARRAYSIZE( InfoValue ) - 1 ] = PXS_CHAR_NULL;
    m_DriverVersion = InfoValue;
    m_DriverVersion.Trim();

    // ODBC Version.
    memset( InfoValue, 0, sizeof ( InfoValue ) );
    StringLength = 0;
    m_ODBC.GetInfo( m_ConnectionHandle,
                    SQL_ODBC_VER, InfoValue, ARRAYSIZE( InfoValue ), &StringLength );
    InfoValue[ ARRAYSIZE( InfoValue ) - 1 ] = PXS_CHAR_NULL;
    m_OdbcVersion = InfoValue;
    m_OdbcVersion.Trim();

    // Data Source Name
    memset( InfoValue, 0, sizeof ( InfoValue ) );
    StringLength = 0;
    m_ODBC.GetInfo( m_ConnectionHandle,
                    SQL_DATA_SOURCE_NAME, InfoValue, ARRAYSIZE( InfoValue ), &StringLength );
    InfoValue[ ARRAYSIZE( InfoValue ) - 1 ] = PXS_CHAR_NULL;
    m_DataSourceName = InfoValue;
    m_DataSourceName.Trim();

    // Server name
    memset( InfoValue, 0, sizeof ( InfoValue ) );
    StringLength = 0;
    m_ODBC.GetInfo( m_ConnectionHandle,
                    SQL_SERVER_NAME, InfoValue, ARRAYSIZE( InfoValue ), &StringLength );
    InfoValue[ ARRAYSIZE( InfoValue ) - 1 ] = PXS_CHAR_NULL;
    m_ServerName = InfoValue;
    m_ServerName.Trim();

    // Database name - not storing at class level because may have only
    // connected to a server and awaiting 'USE <database>'
    memset( InfoValue, 0, sizeof ( InfoValue ) );
    StringLength = 0;
    m_ODBC.GetInfo( m_ConnectionHandle,
                    SQL_DATABASE_NAME, InfoValue, ARRAYSIZE( InfoValue ), &StringLength );
    InfoValue[ ARRAYSIZE( InfoValue ) - 1 ] = PXS_CHAR_NULL;
    DatabaseName = InfoValue;
    DatabaseName.Trim();

    // User name
    memset( InfoValue, 0, sizeof ( InfoValue ) );
    StringLength = 0;
    m_ODBC.GetInfo( m_ConnectionHandle,
                    SQL_USER_NAME, InfoValue, ARRAYSIZE( InfoValue ), &StringLength );
    InfoValue[ ARRAYSIZE( InfoValue ) - 1 ] = PXS_CHAR_NULL;
    m_UserName = InfoValue;
    m_UserName.Trim();

    ////////////////////////////////////////////////////////////////////////////
    // Override values obtained from SQLGetInfo

    // MySQL
    if ( m_DbmsName.CompareI( PXS_DBMS_NAME_MYSQL ) == 0 )
    {
        // The server name includes the protocol, if on the localhost, set that
        if ( m_ServerName.CompareI( L"localhost" ) == 0 )
        {
            m_ServerName = L"localhost";
        }
    }

    // Log the database name etc.
    PXSLogAppInfo1( L"DBMS Name       : '%%1'.", m_DbmsName       );
    PXSLogAppInfo1( L"DBMS Version    : '%%1'.", m_DbmsVersion    );
    PXSLogAppInfo1( L"Driver Mgr. Ver.: '%%1'.", m_DmVersion      );
    PXSLogAppInfo1( L"Driver Name     : '%%1'.", m_DriverName     );
    PXSLogAppInfo1( L"Driver Version  : '%%1'.", m_DriverVersion  );
    PXSLogAppInfo1( L"ODBC Version    : '%%1'.", m_OdbcVersion    );
    PXSLogAppInfo1( L"Data Source Name: '%%1'.", m_DataSourceName );
    PXSLogAppInfo1( L"Server Name     : '%%1'.", m_ServerName     );
    PXSLogAppInfo1( L"Database Name   : '%%1'.", DatabaseName  );
    PXSLogAppInfo1( L"User Name       : '%%1'.", m_UserName       );
}

//===============================================================================================//
//  Description:
//      Disconnect from the database and release and used resources
//
//  Parameters:
//      None
//
//  Remarks:
//      Must not throw, called by destructor
//      Even if the class scope flag m_bConnected is false, still execute this
//      method because want to undo the resource allocation that takes place
//      in the Connect method.
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::Disconnect()
{
    if ( m_bConnected )
    {
        PXSLogAppInfo( L"Disconnecting from database." );
    }

    // Reset class scope data
    m_DbmsName       = PXS_STRING_EMPTY;
    m_DbmsVersion    = PXS_STRING_EMPTY;
    m_DmVersion      = PXS_STRING_EMPTY;
    m_DriverName     = PXS_STRING_EMPTY;
    m_DriverVersion  = PXS_STRING_EMPTY;
    m_OdbcVersion    = PXS_STRING_EMPTY;
    m_DataSourceName = PXS_STRING_EMPTY;
    m_ServerName     = PXS_STRING_EMPTY;
    m_UserName       = PXS_STRING_EMPTY;

    RollbackTrans();
    if ( m_hTransactionStatement )
    {
        m_ODBC.FreeHandle( SQL_HANDLE_STMT, m_hTransactionStatement );
        m_hTransactionStatement = nullptr;
    }

    // Disconnect, must do first before releasing the handle
    if ( m_ConnectionHandle )
    {
        // Ensure connected otherwise risk error 08003
        if ( m_bConnected )
        {
            m_ODBC.Disconnect( m_ConnectionHandle );
        }
        m_ODBC.FreeHandle( SQL_HANDLE_DBC, m_ConnectionHandle );
        m_ConnectionHandle = nullptr;
    }

    // Environment handle
    if ( m_EnvironmentHandle  )
    {
        m_ODBC.FreeHandle( SQL_HANDLE_ENV, m_EnvironmentHandle );
        m_EnvironmentHandle = nullptr;
    }

    m_bConnected = false;
}

//===============================================================================================//
//  Description:
//      Execute an SQL statement
//
//  Parameters:
//      SqlQuery - the SQL statement
//
//  Remarks
//
//  Returns:
//      count of rows affected, -1 if unknown
//===============================================================================================//
SQLLEN OdbcDatabase::ExecuteDirect( const String& SqlQuery )
{
    SQLLEN    rowCount  = 0;
    SQLHSTMT  StatementHandle = nullptr;
    Formatter Format;

    // Verify connected to database
    if ( IsConnected() == false )
    {
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_NOT_CONNECTED, PXS_STRING_EMPTY, __FUNCTION__ );
    }

    if ( SqlQuery.IsEmpty() )
    {
        throw ParameterException( L"SqlQuery", __FUNCTION__ );
    }

    m_ODBC.AllocHandle( SQL_HANDLE_STMT, m_ConnectionHandle, &StatementHandle );
    try
    {
        SetQueryTimeoutAttr( StatementHandle );
        m_ODBC.ExecDirect( StatementHandle, SqlQuery );
        m_ODBC.RowCount( StatementHandle, &rowCount );

        // Log the row count, ignore INSERT as can only insert
        // one record at a time
        if ( ( rowCount > 0 ) &&
             ( SqlQuery.StartsWith( L"INSERT", false ) ) == false )
        {
            PXSLogAppInfo1( L"Affected rows = %%1.", Format.Int64( rowCount ) );
        }
    }
    catch ( const Exception& )
    {
        m_ODBC.FreeHandle( SQL_HANDLE_STMT, StatementHandle );
        throw;
    }
    m_ODBC.FreeHandle( SQL_HANDLE_STMT, StatementHandle );

    return rowCount;
}

//===============================================================================================//
//  Description:
//      Execute a SELECT query and return the SQLINTEGER result
//
//  Parameters:
//      SqlQuery    - the SQL statement
//      pSqlInteger - receives the integer value of a SELECT query
//
//  Returns:
//      true if found a record, else false
//===============================================================================================//
bool OdbcDatabase::ExecuteSelectSqlInteger( const String& SqlQuery, SQLINTEGER* pSqlInteger )
{
    bool          success = false;
    String        FieldValue;
    Formatter     Format;
    OdbcRecordSet RecordSet;

    if ( pSqlInteger == nullptr )
    {
        throw ParameterException( L"pSqlInteger", __FUNCTION__ );
    }
    *pSqlInteger = 0;

    // Only expecting 1 value
    RecordSet.Open( SqlQuery,  this );
    if ( RecordSet.Move( 0 ) == false )
    {
        return false;   // No data
    }

    FieldValue = RecordSet.FieldValue( 0 );
    if ( FieldValue.GetLength() )
    {
        *pSqlInteger = Format.StringToInt32( FieldValue );
        success = true;
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Execute an SQL statement within an existing transaction
//
//  Parameters:
//      pszSQL - pointer to the SQL statement
//      maxRecordsAffected - maximum allowed affected rows, default is -1
//
//  Remarks:
//      If a roll back occurs then any other statements that might have been
//      executed within the transaction will also be rolled back
//
//  Returns:
//      Void
//===============================================================================================//
void OdbcDatabase::ExecuteTrans( const String& SqlQuery, SQLLEN maxRecordsAffected )
{
    SQLLEN rowCount = 0;
    String ErrorMessage;

    // Verify connected to database
    if ( IsConnected() == false )
    {
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_NOT_CONNECTED,
                         PXS_STRING_EMPTY, __FUNCTION__ );
    }

    if ( SqlQuery.IsEmpty() )
    {
        throw ParameterException( L"SqlQuery", __FUNCTION__ );
    }

    // Verify have a statement handle
    if ( m_hTransactionStatement == nullptr )
    {
        throw FunctionException( L"m_hTransactionStatement", __FUNCTION__ );
    }

    // Execute it
    FreeTransactionStatement();
    SetQueryTimeoutAttr( m_hTransactionStatement );
    m_ODBC.ExecDirect( m_hTransactionStatement, SqlQuery );
    if ( PXS_SQLLEN_MAX != maxRecordsAffected )
    {
        // Get the count, if it is unavailable then the returned value is -1,
        // for example in a SELECT query
        m_ODBC.RowCount( m_hTransactionStatement, &rowCount );
        if ( ( rowCount != -1 ) &&
             ( rowCount > maxRecordsAffected ) )
        {
            FreeTransactionStatement();     // Reuse the statement
            RollbackTrans();
            PXSGetResourceString( PXS_IDS_127_TOO_MANY_ROWS_AFFECT,
                                             &ErrorMessage );
            ErrorMessage += PXS_STRING_CRLF;
            ErrorMessage += L"SQL:\r\n";
            ErrorMessage += SqlQuery;
            ErrorMessage += PXS_STRING_CRLF;
            throw Exception( PXS_ERROR_TYPE_APPLICATION,
                             PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), __FUNCTION__ );
        }
    }
    FreeTransactionStatement();     // Reuse the statement
}

//===============================================================================================//
//  Description:
//     Get the database's connection handle.
//
//  Parameters:
//      None
//
//  Returns:
//      the SQLHDBC handle, NULL if not connected
//
//===============================================================================================//
SQLHDBC OdbcDatabase::GetConnectionHandle() const
{
    return m_ConnectionHandle;
}

//===============================================================================================//
//  Description:
//      Clean a string for an SQL statement
//
//  Parameters:
//      Input   - pointer to input string
//      sqlType - named constant of the SQL data type e.g. SQL_INTEGER
//      pOutput - string object to receive the output
//
//  Remarks:
//      Output string receives "NULL" if the input is null or empty.
//      Could use SQLGetInfo but will use the known string literals
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::FixUpStringSQL( const String& Input,
                                   SQLSMALLINT sqlType, String* pOutput ) const
{
    String Clean;

    if ( pOutput == nullptr )
    {
        throw ParameterException( L"pOutput", __FUNCTION__ );
    }

    // Test for NULL/Empty
    if ( Input.IsEmpty() )
    {
        *pOutput = L"NULL";
        return;
    }
    *pOutput = PXS_STRING_EMPTY;

    // Escape single quotes by using two single quotes.
    // Use 0x27 = apostrophe as cccc fails if use '\''
    Clean = Input;
    Clean.ReplaceChar( 0x27, L"''" );

    // Escape strings and dates/times
    if ( IsStringDataType( sqlType ) )
    {
        if ( m_DbmsName.CompareI( PXS_DBMS_NAME_MYSQL ) == 0 )
        {
            // mysql_real_escape_string says to backslashes
            Clean.ReplaceChar( '\\', L"\\\\" );
        }
        else if ( ( Clean.IsOnlyUSAscii() ) == false &&
                  ( m_DbmsName.CompareI( PXS_DBMS_NAME_SQL_SERVER ) ) == 0 )
        {
            // High ASCII/Unicode so prefix with N for SQL Server
            *pOutput += L"N";
        }

        *pOutput += L"'";
        *pOutput += Clean;
        *pOutput += L"'";
    }
    else if ( IsDateDataType( sqlType ) )
    {
        if ( m_DbmsName.CompareI( PXS_DBMS_NAME_ACCESS ) == 0 )
        {
            // Access uses # to delimit dates
            *pOutput  = L"#";
            *pOutput += Clean;
            *pOutput += L"#";
        }
        else if ( m_DbmsName.IndexOfI( PXS_DBMS_NAME_ORACLE )  != PXS_MINUS_ONE )
        {
            if ( sqlType == SQL_TYPE_DATE )
            {
                *pOutput = L"{ d '";       // { d 'YYYY:MM:DD' }
            }
            else if ( sqlType == SQL_TYPE_TIME )
            {
                *pOutput = L"{ t '";       // { t 'hh:mm:ss' }
            }
            else if ( sqlType == SQL_TYPE_TIMESTAMP )
            {
                *pOutput = L"{ ts '";      // { ts 'YYYY-MM-DD hh:mm:ss' }
            }
            *pOutput += Clean;
            *pOutput += L"' }";
        }
        else
        {
            // Default is to quote with '
            *pOutput  = L"'";
            *pOutput += Clean;
            *pOutput += L"'";
        }
    }
    else
    {
        // Number etc. so use the input value
        *pOutput = Clean;
    }
}

//===============================================================================================//
//  Description:
//     Get the name of the database
//
//  Parameters:
//      pDatabaseName - receives the database name
//
//  Remarks:
//      The database name is not stored at class scope after connecting because
//      may have connected to the server only, then subsequently executed
//      'USE <database>' so need to dynamically query for the database's name.
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::GetDatabaseName( String* pDatabaseName )
{
    wchar_t     szDatabaseName[ MAX_PATH + 1 ] = { 0 };
    SQLSMALLINT StringLength = 0;

    if ( pDatabaseName == nullptr )
    {
        throw ParameterException( L"pDatabaseName", __FUNCTION__ );
    }
    *pDatabaseName = PXS_STRING_EMPTY;

    if ( IsConnected() == false )
    {
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_NOT_CONNECTED, PXS_STRING_EMPTY, __FUNCTION__ );
    }

    // Get the database name, must have this for the catalogue
    m_ODBC.GetInfo( m_ConnectionHandle,
                    SQL_DATABASE_NAME,
                    szDatabaseName, ARRAYSIZE( szDatabaseName ), &StringLength );
    szDatabaseName[ ARRAYSIZE( szDatabaseName ) - 1 ] = PXS_CHAR_NULL;
    *pDatabaseName = szDatabaseName;
    pDatabaseName->Trim();

    // The access driver 'Microsoft Access Driver (*.mdb)' does not append
    // the file extension so add it if its missing. The newer driver
    // (*.mdb, *.accdb) returns the database name with the file extension
    // so would never need to add .accdb
    if ( ( pDatabaseName->GetLength() ) &&
         ( m_DbmsName.CompareI( PXS_DBMS_NAME_ACCESS ) ) == 0 )
    {
        if ( ( pDatabaseName->EndsWithStringI( L".mdb"   ) == false ) &&
             ( pDatabaseName->EndsWithStringI( L".accdb" ) == false )  )
        {
            *pDatabaseName += L".mdb";
        }
    }
}

//===============================================================================================//
//  Description:
//     Get a keyword for a given database
//
//  Parameters:
//      pszAnsiKeyword - the ANSI keyword, e.g. USER
//      pDbKeyword     - receives the database specific keyword
//
//  Remarks:
//      E.g. Access uses CurrentUser() instead of CURRENT_USER
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::GetDbKeyWord( LPCWSTR pszAnsiKeyword, String* pDbKeyword )
{
    String AnsiKeyword;

    if ( pDbKeyword == nullptr )
    {
        throw ParameterException( L"pDbKeyword", __FUNCTION__ );
    }
    *pDbKeyword = PXS_STRING_EMPTY;

    // Must be connected
    if ( IsConnected() == false )
    {
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_NOT_CONNECTED, PXS_STRING_EMPTY, __FUNCTION__ );
    }

    AnsiKeyword = pszAnsiKeyword;
    AnsiKeyword.Trim();
    if ( AnsiKeyword.CompareI( PXS_KEYWORD_CURRENT_TIMESTAMP ) == 0 )
    {
        GetKeywordCurrentTimestamp( pDbKeyword );
    }
    else if ( AnsiKeyword.CompareI( PXS_KEYWORD_UPPER ) == 0 )
    {
        GetKeywordUpper( pDbKeyword );
    }
    else if ( AnsiKeyword.CompareI( PXS_KEYWORD_TIMESTAMP ) == 0 )
    {
        GetKeywordUtcTimestamp( pDbKeyword );
    }
    else
    {
        throw ParameterException( L"pszAnsiKeyword", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//     Get the name of the database management system.
//
//  Parameters:
//      None
//
//  Remarks:
//      Must be connected or get an empty string
//
//  Returns:
//      Constant reference to the database management system name
//===============================================================================================//
const String& OdbcDatabase::GetDbmsName() const
{
    return m_DbmsName;
}

//===============================================================================================//
//  Description:
//     Get the database's major and minor version numbers
//
//  Parameters:
//      pMajor - receives the major version
//      pMinor - receives the minor version
//
//  Returns:
//      -1 for both values if cannot determine the versions
//===============================================================================================//
void OdbcDatabase::GetMajorMinorVersion( int* pMajor, int* pMinor ) const
{
    String      Token;
    Formatter   Format;
    StringArray Tokens;

    if ( ( pMajor == nullptr ) || ( pMinor == nullptr ) )
    {
        throw ParameterException( L"pMajor/pMinor", __FUNCTION__ );
    }
    *pMajor = -1;
    *pMinor = -1;

    // Explode the version string
    if ( m_DbmsVersion.GetLength() )
    {
        m_DbmsVersion.ToArray( PXS_CHAR_DOT, &Tokens );
        if ( Tokens.GetSize() > 1 )
        {
            Token = Tokens.Get( 0 );
            if ( Token.c_str() )
            {
                *pMajor = Format.StringToInt32( Token );
            }

            Token = Tokens.Get( 1 );
            if ( Token.c_str() )
            {
                *pMinor = Format.StringToInt32( Token );
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Get a list of the procedures in the database
//
//  Parameters:
//      Filter      - filter to reduce results
//      pProcedures - string array to receive the procedure names
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::GetProcedures( const String& Filter, StringArray* pProcedures )
{
    size_t   idxSemiColon = 0;
    wchar_t  szProcName[ 256 ] = { 0 };     // Big enough for any procedure name
    SQLLEN   StrLen_or_Ind = 0;
    String   ProcName, FilterClean, DatabaseName, NullSchemaName;
    SQLHSTMT StatementHandle = nullptr;
    SQLUSMALLINT sqlIdentifierCase = 0;

    if ( pProcedures == nullptr )
    {
        throw ParameterException( L"pProcedures", __FUNCTION__ );
    }
    pProcedures->RemoveAll();

    if ( SupportsStoredProcedures() == false )
    {
        return;     // Nothing to do
    }

    // Must have a database name as the catalogue
    GetDatabaseName( &DatabaseName );
    if ( DatabaseName.IsEmpty() )
    {
        throw FunctionException( L"DatabaseName", __FUNCTION__ );
    }

    // Adjust case of the filter, if required
    FilterClean = Filter;
    sqlIdentifierCase = SQL_IC_MIXED;
    m_ODBC.GetInfo( m_ConnectionHandle,
                    SQL_IDENTIFIER_CASE,
                    &sqlIdentifierCase, sizeof ( sqlIdentifierCase ), nullptr );
    if ( sqlIdentifierCase == SQL_IC_UPPER )
    {
        FilterClean.ToUpperCase();
    }
    else if ( sqlIdentifierCase == SQL_IC_LOWER )
    {
        FilterClean.ToLowerCase();
    }

    // Catch all exceptions so can clean up the statement handle
    m_ODBC.AllocHandle( SQL_HANDLE_STMT, m_ConnectionHandle, &StatementHandle);
    try
    {
        SetQueryTimeoutAttr( StatementHandle );
        m_ODBC.Procedures( StatementHandle,
                           DatabaseName,
                           NullSchemaName,      // All schemas
                           FilterClean );

        m_ODBC.BindCol( StatementHandle,
                        3,                     // = PROCEDURE_NAME
                        SQL_C_TCHAR, szProcName, ARRAYSIZE( szProcName ), &StrLen_or_Ind );

        memset( szProcName, 0, sizeof ( szProcName ) );
        while ( SQL_NO_DATA != m_ODBC.Fetch( StatementHandle ) )
        {
            szProcName[ ARRAYSIZE( szProcName ) - 1 ] = PXS_CHAR_NULL;

            // Clean the procedure name. SQL Server has a semi-colon in the name
            ProcName = szProcName;
            ProcName.Trim();
            idxSemiColon = ProcName.IndexOf( ';', 0 );
            if ( idxSemiColon < ProcName.GetLength()  )
            {
                ProcName.Truncate( idxSemiColon );
            }
            ProcName.Trim();

            if ( ProcName.GetLength() )
            {
                pProcedures->Add( ProcName );
            }
            memset( szProcName, 0, sizeof ( szProcName ) );  // Next pass
        }
    }
    catch ( const Exception& )
    {
        m_ODBC.FreeHandle( SQL_HANDLE_STMT, StatementHandle );
        throw;
    }
    m_ODBC.FreeHandle( SQL_HANDLE_STMT, StatementHandle );
}

//===============================================================================================//
//  Description:
//     Get the name of the server on which the database is located
//
//  Parameters:
//      None
//
//  Remarks:
//      Must be connected or the returned string will be
//
//  Returns:
//      Constant reference to the name of the server
//===============================================================================================//
const String& OdbcDatabase::GetServerName() const
{
    return m_ServerName;
}

//===============================================================================================//
//  Description:
//     Get if the database supports a specified data type
//
//  Parameters:
//      sqlType   - the SQL data type, eg. SQL_INTEGER
//      pTypeName - receives the name of the type
//
//  Remarks:
//      SQLGetTypeInfo causes a round trip to the server
//
//  Returns:
//      true of the data type is supported otherwise false
//===============================================================================================//
bool OdbcDatabase::GetSupportsSqlType( SQLSMALLINT sqlType, String* pTypeName )
{
    bool        supported  = false;
    wchar_t     szTypeName[ 256 ] = { 0 };      // Big enough for any type name
    SQLLEN      StrLen_or_IndType = 0, StrLen_or_IndName = 0;
    SQLHSTMT    StatementHandle = nullptr;
    SQLSMALLINT dataType = 0;

    if ( pTypeName == nullptr )
    {
        throw ParameterException( L"pTypeName", __FUNCTION__ );
    }
    *pTypeName = PXS_STRING_EMPTY;

    if ( IsConnected() == false )
    {
        throw Exception ( PXS_ERROR_TYPE_APPLICATION,
                          PXS_ERROR_DB_NOT_CONNECTED, PXS_STRING_EMPTY, __FUNCTION__ );
    }

    // Catch all exceptions so can clean up the statement handle
    m_ODBC.AllocHandle( SQL_HANDLE_STMT, m_ConnectionHandle, &StatementHandle);
    try
    {
        SetQueryTimeoutAttr( StatementHandle );
        m_ODBC.GetTypeInfo( StatementHandle, sqlType );
        m_ODBC.BindCol( StatementHandle,
                        1,             // = TYPE_NAME
                        SQL_C_TCHAR, szTypeName, ARRAYSIZE( szTypeName ), &StrLen_or_IndName );

        // Bind the data type
        StrLen_or_IndType = 0;
        m_ODBC.BindCol( StatementHandle,
                        2,             // = DATA_TYPE
                        SQL_C_SSHORT, &dataType, 0, &StrLen_or_IndType );

        // Fetch the rows, there may be more than one. The driver will map the
        // requested type to the ones it supports or database domains
        supported = false;
        while ( ( supported == false ) &&
                ( SQL_NO_DATA != m_ODBC.Fetch( StatementHandle ) ) )
        {
            szTypeName[ ARRAYSIZE( szTypeName ) - 1 ] = PXS_CHAR_NULL;
            if ( dataType == sqlType )
            {
                supported  = true;
                *pTypeName = szTypeName;
            }
        }
    }
    catch ( const Exception& )
    {
        m_ODBC.FreeHandle( SQL_HANDLE_STMT, StatementHandle );
        throw;
    }
     m_ODBC.FreeHandle( SQL_HANDLE_STMT, StatementHandle );

    return supported;
}

//===============================================================================================//
//  Description:
//     Get a list of tables in the database
//
//  Parameters:
//      wantUser   - when true indicates want user tables
//      wantSystem - when true indicates want system tables
//      wantViews  - when true indicates want views
//      pTables    - array object to receive the table names
//
//  Remarks:
//      Array is sorted
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::GetTables( bool wantUser,
                              bool wantSystem, bool wantViews, StringArray* pTables )
{
    wchar_t  szTableName[ 256 ] = { 0 };  // Enough for any table name
    SQLLEN   StrLen_or_Ind = 0;
    String   TableType, DatabaseName, NullSchemaName, NullTableName;
    SQLHSTMT StatementHandle = nullptr;

    if ( pTables == nullptr )
    {
        throw ParameterException( L"pTables", __FUNCTION__ );
    }
    pTables->RemoveAll();

    if ( IsConnected() == false )
    {
        throw Exception ( PXS_ERROR_TYPE_APPLICATION,
                          PXS_ERROR_DB_NOT_CONNECTED, PXS_STRING_EMPTY, __FUNCTION__ );
    }

    // Must have a database name as the catalogue
    GetDatabaseName( &DatabaseName );
    if ( DatabaseName.IsEmpty() )
    {
       throw FunctionException( L"DatabaseName", __FUNCTION__ );
    }

    // Copy in the table type(s), these should be comma separated
    if ( wantUser )
    {
        TableType = L"TABLE";
    }

    if ( wantSystem )
    {
        if ( TableType.GetLength() )
        {
            TableType += L", ";
        }
        TableType += L"SYSTEM TABLE";
    }

    if ( wantViews )
    {
        if ( TableType.GetLength() )
        {
            TableType += L", ";
        }
        TableType += L"VIEW";
    }

    // Catch all exceptions so can clean up the statement handle
    m_ODBC.AllocHandle( SQL_HANDLE_STMT, m_ConnectionHandle, &StatementHandle);
    try
    {
        SetQueryTimeoutAttr( StatementHandle );
        m_ODBC.Tables( StatementHandle,
                       DatabaseName,
                       NullSchemaName,          // All schemas
                       NullTableName,           // All tables
                       TableType );

        m_ODBC.BindCol( StatementHandle,
                        3,             // = TABLE_NAME
                        SQL_C_TCHAR, szTableName, ARRAYSIZE( szTableName ), &StrLen_or_Ind );

        memset( szTableName, 0, sizeof ( szTableName ) );
        while ( SQL_NO_DATA != m_ODBC.Fetch( StatementHandle ) )
        {
            szTableName[ ARRAYSIZE( szTableName ) - 1 ] = PXS_CHAR_NULL;
            pTables->Add( szTableName );
            memset( szTableName, 0, sizeof ( szTableName ) );   // Next pass
        }
    }
    catch ( const Exception& )
    {
        m_ODBC.FreeHandle( SQL_HANDLE_STMT, StatementHandle );
        throw;
    }
    m_ODBC.FreeHandle( SQL_HANDLE_STMT, StatementHandle );
}

//===============================================================================================//
//  Description:
//     Get the name of the logged on user
//
//  Parameters:
//      None
//
//  Remarks:
//      Must be connected or the returned string will be empty
//
//  Returns:
//      Constant reference to the name of the logged on user
//===============================================================================================//
const String& OdbcDatabase::GetUserName() const
{
    return m_UserName;
}

//===============================================================================================//
//  Description:
//      Get if the default behaviour of database management system is
//      to perform a case sensitive sorting and grouping
//
//  Parameters:
//      pszDbmsName  - the database management system name
//
//  Remarks:
//      Determined as documented by the DBMS vendor
//      Access      = False
//      MySQL       = False
//      Oracle      = True
//      SQL Server  = False, but can be installed to be case sensitive
//
//  Returns:
//      true if database does case sensitive sorting and grouping else false
//===============================================================================================//
bool OdbcDatabase::IsCaseSensitiveSort() const
{
    // Need a DBMS name
    if ( m_DbmsName.IsEmpty() )
    {
        throw FunctionException( L"m_DbmsName", __FUNCTION__ );
    }

    // The follow are do case sensitive sorting and grouping
    if ( ( m_DbmsName.CompareI( PXS_DBMS_NAME_ORACLE      ) == 0 ) ||
         ( m_DbmsName.CompareI( PXS_DBMS_NAME_POSTGRE_SQL ) == 0 )  )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      See if connected to a database
//
//  Parameters:
//      None
//
//  Returns:
//      true if connected, else false
//===============================================================================================//
bool OdbcDatabase::IsConnected() const
{
    return m_bConnected;
}

//===============================================================================================//
//  Description:
//      Set the time out for queries
//
//  Parameters:
//      queryTimeOut - the time out in seconds
//
//  Remarks:
//      Zero means no time out
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::QueryTimeOut( DWORD queryTimeOut )
{
    m_uQueryTimeOut = queryTimeOut;
}

//===============================================================================================//
//  Description:
//      Get the time out for queries
//
//  Parameters:
//      None
//
//  Remarks:
//      Zero means no time out
//
//  Returns:
//      the query time out in seconds
//===============================================================================================//
DWORD OdbcDatabase::QueryTimeOut() const
{
    return m_uQueryTimeOut;
}

//===============================================================================================//
//  Description:
//      Roll back an open transaction
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::RollbackTrans()
{
    Formatter Format;

    // Check if connected to database
    if ( IsConnected() == false )
    {
        return;     // Nothing to do
    }

    if ( false == StartedTrans() )
    {
        return;     // Nothing to do
    }
    FreeTransactionStatement();     // Reuse the statement

    // Roll back
    m_ODBC.EndTran( SQL_HANDLE_DBC, m_ConnectionHandle, SQL_ROLLBACK );
    m_ODBC.SetConnectAttr( m_ConnectionHandle,
                           SQL_ATTR_AUTOCOMMIT,
                           reinterpret_cast<SQLPOINTER>( SQL_AUTOCOMMIT_ON ), SQL_IS_UINTEGER );

    PXSLogAppInfo( L"A transaction was rolled back." );
}

//===============================================================================================//
//  Description:
//      Determine if a transaction has been started
//
//  Parameters:
//      None
//
//  Returns:
//      true if a transaction has been started, otherwise false
//===============================================================================================//
bool OdbcDatabase::StartedTrans()
{
    bool        started = false;
    SQLINTEGER  StringLength = 0;
    SQLUINTEGER autoCommit   = 0;

    // See if connected to database
    if ( IsConnected() == false )
    {
        return false;
    }

    // Get the auto commit value
    m_ODBC.GetConnectAttr( m_ConnectionHandle,
                           SQL_ATTR_AUTOCOMMIT, &autoCommit, SQL_IS_UINTEGER, &StringLength );

    // If the auto commit is off then changes are only committed
    // when SQLEndTran is called, i.e. a transaction has been started
    if ( autoCommit == SQL_AUTOCOMMIT_OFF )
    {
        started = true;
    }

    return started;
}

//===============================================================================================//
//  Description:
//      Determine if the database supports the query time out feature
//
//  Parameters:
//      None
//
//  Remarks:
//      Specifying a query time out using SQLSetStmtAttr(SQL_ATTR_QUERY_TIMEOUT)
//      may result in an error if this feature is not supported. Unfortunately,
//      SQLGetInfo does not report the time out feature so will use documented
//      behaviour.
//
//      Default is false.
//
//      Access does not support a time out, looks like PostgreSQL does
//      not either
//
//  Returns:
//      true if supports a query time out feature
//===============================================================================================//
bool OdbcDatabase::SupportsQueryTimeOut() const
{
    // Compare with the DBMS name
    if ( ( m_DbmsName.CompareI( PXS_DBMS_NAME_MYSQL      ) == 0 ) ||
         ( m_DbmsName.CompareI( PXS_DBMS_NAME_ORACLE     ) == 0 ) ||
         ( m_DbmsName.CompareI( PXS_DBMS_NAME_SQL_SERVER ) == 0 )  )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Determine if the database supports stored procedures
//
//  Parameters:
//      None
//
//  Remarks:
//      Versions prior to Access 2000 and MySQL 5.0 do not support procedures
//
//  Returns:
//      true if a supports stored procedures
//===============================================================================================//
bool OdbcDatabase::SupportsStoredProcedures()
{
    bool  supported = false;
    int   major = 0, minor   = 0;
    SQLWCHAR    InfoValue[ 64 ]    = { 0 };  // N.B. must an even number
    SQLSMALLINT StringLength = 0;

    if ( m_ConnectionHandle == nullptr )
    {
        throw FunctionException( L"m_ConnectionHandle", __FUNCTION__ );
    }

    if ( m_DbmsName.IsEmpty() )
    {
        throw FunctionException( L"m_DbmsName", __FUNCTION__ );
    }

    m_ODBC.GetInfo( m_ConnectionHandle,
                    SQL_PROCEDURES, InfoValue, ARRAYSIZE( InfoValue ), &StringLength );
    InfoValue[ ARRAYSIZE( InfoValue ) - 1 ] = PXS_CHAR_NULL;

    // The result is either Y or N
    if ( 'Y' == toupper( InfoValue[ 0 ] ) )
    {
        supported = true;
    }

    ////////////////////////////////////////////////////////////////////////////
    // Override values obtained from SQLGetInfo

    // Prior to JET v4 procedures not supported
    if ( m_DbmsName.CompareI( PXS_DBMS_NAME_ACCESS ) == 0 )
    {
        GetMajorMinorVersion( &major, &minor );
        if ( major < 4 )
        {
            supported = false;
        }
    }

    // MySQL Driver v5.0+ supports procedures
    if ( m_DbmsName.CompareI( PXS_DBMS_NAME_MYSQL ) == 0 )
    {
        GetMajorMinorVersion( &major, &minor );
        if ( major >= 5 )
        {
            supported = true;
        }
    }

    // PostgreSQL does not support CREATE PROCEDURE
    if ( m_DbmsName.CompareI( PXS_DBMS_NAME_POSTGRE_SQL ) == 0 )
    {
        supported = false;
    }

    return supported;
}

//===============================================================================================//
//  Description:
//      Determine if the database supports views
//
//  Parameters:
//      None
//
//  Remarks:
//      Versions prior to Access 2000 and MySQL 5.0 do not support views
//
//  Returns:
//      true if a supports views, otherwise false
//===============================================================================================//
bool OdbcDatabase::SupportsViews()
{
    bool supported = false;
    int  major = 0, minor = 0;
    SQLUINTEGER InfoValue = 0;

    if ( m_ConnectionHandle == nullptr )
    {
        throw FunctionException( L"m_ConnectionHandle", __FUNCTION__ );
    }

    if ( m_DbmsName.GetLength() )
    {
        throw FunctionException( L"m_DbmsName", __FUNCTION__ );
    }

    // Views
    // If using Oracle with the Microsoft ODBC for Oracle, get the message
    // 'Information type out of range Native Error Code: State: HY096'. This
    // does not happen with the Oracle driver.
    try
    {
        m_ODBC.GetInfo( m_ConnectionHandle,
                        SQL_CREATE_VIEW, &InfoValue, sizeof ( InfoValue ), nullptr );
        if ( InfoValue )
        {
            supported = true;
        }
    }
    catch ( const Exception& e )
    {
        // Log and continue
        PXSLogException( L"Error determining if views are supported",
                         e, __FUNCTION__ );
    }

    ////////////////////////////////////////////////////////////////////////////
    // Override values obtained from SQLGetInfo

    // Prior JET v4 and views not supported
    if ( m_DbmsName.CompareI( PXS_DBMS_NAME_ACCESS ) == 0 )
    {
        GetMajorMinorVersion( &major, &minor );
        if ( major < 4 )
        {
            supported = false;
        }
    }

    // MySQL Driver v5.0+ supports views
    if ( m_DbmsName.CompareI( PXS_DBMS_NAME_MYSQL ) == 0 )
    {
        GetMajorMinorVersion( &major, &minor );
        if ( major >= 5 )
        {
            supported = true;
        }
    }

    return supported;
}


///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//     // Free the statement so can re-use it
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::FreeTransactionStatement()
{
    if ( m_hTransactionStatement )
    {
        m_ODBC.FreeStmt( m_hTransactionStatement, SQL_UNBIND );
        m_ODBC.FreeStmt( m_hTransactionStatement, SQL_RESET_PARAMS );
        m_ODBC.FreeStmt( m_hTransactionStatement, SQL_CLOSE );
    }
}

//===============================================================================================//
//  Description:
//     Get the keyword used by the database for CURRENT_TIMESTAMP
//
//  Parameters:
//      pDbKeyword - receives the keyword
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::GetKeywordCurrentTimestamp( String* pDbKeyword )
{
    if ( pDbKeyword == nullptr )
    {
        throw ParameterException( L"pDbKeyword", __FUNCTION__ );
    }

    // Need a DBMS name
    if ( m_DbmsName.IsEmpty() )
    {
        throw FunctionException( L"m_DbmsName", __FUNCTION__ );
    }

    if ( m_DbmsName.CompareI( PXS_DBMS_NAME_ACCESS ) == 0 )
    {
        *pDbKeyword = L"NOW";   // Does not seem to need the brackets
    }
    else if ( m_DbmsName.CompareI( PXS_DBMS_NAME_ORACLE ) == 0 )
    {
        *pDbKeyword = L"SYSDATE";
    }
    else
    {
        *pDbKeyword = PXS_KEYWORD_CURRENT_TIMESTAMP;
    }
}

//===============================================================================================//
//  Description:
//     Get the keyword used by the database for PXS_KEYWORD_UPPER
//
//  Parameters:
//      pDbKeyword - receives the keyword
//
//  Returns:
//     void
//===============================================================================================//
void OdbcDatabase::GetKeywordUpper( String* pDbKeyword )
{
    if ( pDbKeyword == nullptr )
    {
        throw ParameterException( L"pDbKeyword", __FUNCTION__ );
    }

    // Need a DBMS name
    if ( m_DbmsName.IsEmpty() )
    {
        throw FunctionException( L"m_DbmsName", __FUNCTION__ );
    }

    if (  m_DbmsName.CompareI( PXS_DBMS_NAME_ACCESS ) == 0 )
    {
        *pDbKeyword = L"UCASE";    // No brackets
    }
    else
    {
        *pDbKeyword = PXS_KEYWORD_UPPER;
    }
}

//===============================================================================================//
//  Description:
//     Get the keyword used by the database for PXS_KEYWORD_TIMESTAMP
//
//  Parameters:
//      pDbKeyword - receives the keyword, "" if not supported
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::GetKeywordUtcTimestamp( String* pDbKeyword )
{
    int major = 0, minor = 0;

    if ( pDbKeyword == nullptr )
    {
        throw ParameterException( L"pDbKeyword", __FUNCTION__ );
    }
    *pDbKeyword = PXS_STRING_EMPTY;

    if ( m_DbmsName.IsEmpty() )
    {
        throw FunctionException( L"m_DbmsName", __FUNCTION__ );
    }

    if ( m_DbmsName.CompareI( PXS_DBMS_NAME_ACCESS ) == 0 )
    {
        *pDbKeyword = PXS_STRING_EMPTY;        // Not supported
    }
    else if ( m_DbmsName.CompareI( PXS_DBMS_NAME_MYSQL ) == 0 )
    {
        // MySQL has UTC_TIMESTAMP() on v4.1.1 and higher.
        GetMajorMinorVersion( &major, &minor );
        if ( ( major == 4 && minor >= 2 ) || ( major >  4 ) )
        {
            *pDbKeyword = L"UTC_TIMESTAMP()";
        }
    }
    else if ( m_DbmsName.CompareI( PXS_DBMS_NAME_ORACLE ) == 0 )
    {
        // For Oracle 9 can use SYS_EXTRACT_UTC( SYSTIMESTAMP ). In a
        // SELECT statement FROM DUAL works fine as a  parameter in
        // an UPDATE statement
        GetMajorMinorVersion( &major, &minor );
        if ( major >= 9 )
        {
            *pDbKeyword = L"SYS_EXTRACT_UTC( SYSTIMESTAMP )";
        }
    }
    else if ( m_DbmsName.CompareI( PXS_DBMS_NAME_SQL_SERVER ) == 0 )
    {
        // SQL Server 2000+
        GetMajorMinorVersion( &major, &minor );
        if ( major >= 8 )
        {
            *pDbKeyword = L"GETUTCDATE()";
        }
    }
    else if ( m_DbmsName.CompareI( PXS_DBMS_NAME_POSTGRE_SQL ) == 0 )
    {
        *pDbKeyword = L"CURRENT_TIMESTAMP(0) AT TIME ZONE 'UTC'";
    }
    else
    {
        *pDbKeyword = PXS_KEYWORD_TIMESTAMP;
    }
}

//===============================================================================================//
//  Description:
//     Test if the specified ODBC data type is a date
//
//  Parameters:
//      sqlType - the data type
//
//  Returns:
//      true if a date type, otherwise false
//===============================================================================================//
bool OdbcDatabase::IsDateDataType( SQLSMALLINT sqlType ) const
{
    if ( ( SQL_TYPE_DATE      == sqlType ) ||    // 91   ODBC 3.0
         ( SQL_TYPE_TIME      == sqlType ) ||    // 92   ODBC 3.0
         ( SQL_TYPE_TIMESTAMP == sqlType )  )    // 93   ODBC 3.0
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//     Test if the specified ODBC data type is a string
//
//  Parameters:
//      sqlType - the data type
//
//  Returns:
//      true if a string type, otherwise false
//===============================================================================================//
bool OdbcDatabase::IsStringDataType( SQLSMALLINT sqlType ) const
{
    if ( ( SQL_CHAR         == sqlType ) ||    // 1
         ( SQL_VARCHAR      == sqlType ) ||    // 12
         ( SQL_LONGVARCHAR  == sqlType ) ||    // (-1)
         ( SQL_WCHAR        == sqlType ) ||    // (-8)
         ( SQL_WVARCHAR     == sqlType ) ||    // (-9)
         ( SQL_WLONGVARCHAR == sqlType )  )    // (-10)
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//     Set the time out for the specified query
//
//  Parameters:
//      StatementHandle - a handle to the statement
//
//  Remarks:
//     Does not throw because setting a time out is an optional feature
//     and want to continue processing the query.
//
//      The time out value is the class scope property m_uQueryTimeOut
//
//  Returns:
//      void
//===============================================================================================//
void OdbcDatabase::SetQueryTimeoutAttr( SQLHSTMT StatementHandle )
{
    if ( StatementHandle == nullptr )
    {
        return;
    }

    try
    {
        if ( SupportsQueryTimeOut() )
        {
            DWORD_PTR dwordPtr = static_cast<DWORD_PTR>( m_uQueryTimeOut );
            m_ODBC.SetStmtAttr( StatementHandle,
                                SQL_ATTR_QUERY_TIMEOUT,
                                reinterpret_cast<SQLPOINTER>( dwordPtr ),
                                SQL_IS_UINTEGER );
        }
    }
    catch ( const Exception& e )
    {
        // Log and continue
        PXSLogException( L"Failed to set query timeout", e, __FUNCTION__ );
    }
}
