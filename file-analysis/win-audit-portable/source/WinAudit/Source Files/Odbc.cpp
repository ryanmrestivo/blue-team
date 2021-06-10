///////////////////////////////////////////////////////////////////////////////////////////////////
//
// ODBC Utility and Wrapper Class Implementation
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
#include "WinAudit/Header Files/Odbc.h"

// 2. C System Files
#include <float.h>
#include <math.h>
#include <sqlext.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateWChars.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/Directory.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/NameValue.h"
#include "PxsBase/Header Files/Registry.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/SystemInformation.h"

// 5. This Project

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
Odbc::Odbc()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
Odbc::~Odbc()
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
//      Get the names of the file DSNs
//
//  Parameters:
//      pDataSourceNames - receives the name of the file DSNs
//
//  Remarks:
//      XP vs Vista
//          - The ODBC admin interface is the same for both
//          - In the registry the ODBC keys are the same for both
//          - CommonFilesDir exists for both at 'HKEY_LOCAL_MACHINE\SOFTWARE
//            \Microsoft\Windows\CurrentVersion'
//          - The directory 'C:\Program Files\Common Files\' exists for both
//          - The key HKEY_CURRENT_USER\Software\ODBC\ODBC.INI\ODBC File DSN'
//            does not exist on Vista
//          - For Vista the default directory for file DSNs is User\Documents
//          - 'ODBC\Data Sources' does not exist below C:\Program Files\Common
//             Files\ on Vista
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::GetFileDataSourceNames( StringArray* pDataSourceNames )
{
    String    DsnDir, DotExtension;
    Registry  RegObject;
    Formatter Format;
    Directory DirObject;

    if ( pDataSourceNames == nullptr )
    {
        throw ParameterException( L"pDataSourceNames", __FUNCTION__ );
    }
    pDataSourceNames->RemoveAll();

    // Get the DefaultDSNDir folder, this does not apply to a Vista fresh
    // install.
    RegObject.Connect( HKEY_CURRENT_USER );
    RegObject.GetStringValue( L"Software\\ODBC\\ODBC.INI\\ODBC File DSN",
                              L"DefaultDSNDir", &DsnDir );
    DsnDir.Trim();

    // If did not get that, try CommonFilesDir
    if ( DsnDir.IsEmpty() )
    {
        RegObject.Connect( HKEY_LOCAL_MACHINE );
        RegObject.GetStringValue(
                        L"SOFTWARE\\Microsoft\\Windows\\CurrentVersion",
                        L"CommonFilesDir", &DsnDir );
        DsnDir.Trim();
        if ( ( DsnDir.GetLength() ) &&
             ( DsnDir.EndsWithCharacter( PXS_PATH_SEPARATOR ) ) == false )
        {
            DsnDir += PXS_PATH_SEPARATOR;
        }

        // This directory may not exist, e.g. Vista
        if ( DsnDir.GetLength() )
        {
            DsnDir += L"ODBC\\Data Sources";
        }
    }

    // Get the file list
    if ( DirObject.Exists( DsnDir ) )
    {
        // Get the files
        DotExtension = L".dsn";
        DirObject.ListFiles( DsnDir, DotExtension, pDataSourceNames );
        PXSLogAppInfo2( L"Found %%1 file DSNs in the directory '%%2'.",
                        Format.SizeT( pDataSourceNames->GetSize() ), DsnDir );
    }
    else
    {
        PXSLogSysWarn1( ERROR_PATH_NOT_FOUND,
                        L"Directory '%%1' not specified or not found.", DsnDir );
    }
    pDataSourceNames->Sort( true );
}

//===============================================================================================//
//  Description:
//      Get the installed drivers
//
//  Parameters:
//      pDrivers - string array to receive the data
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::GetInstalledDrivers( StringArray* pDrivers )
{
    HENV        EnvironmentHandle = nullptr;
    String      ErrorMessage;
    SQLWCHAR    DriverDescription[ 256 ];
    SQLWCHAR    DriverAttributes[ 1024 ];   // Can be long due to file paths
    SQLRETURN   sqlReturn = 0;
    SQLSMALLINT AttributesLengthPtr = 0, DescriptionLengthPtr = 0;

    if ( pDrivers == nullptr )
    {
        throw ParameterException( L"pDrivers", __FUNCTION__ );
    }
    pDrivers->RemoveAll();

    sqlReturn = SQLAllocHandle( SQL_HANDLE_ENV,
                                nullptr,            // = SQL_NULL_HANDLE
                                &EnvironmentHandle );
    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, L"SQLAllocHandle",  __FUNCTION__ );
    }

    // Set to ODBC v3
    sqlReturn = SQLSetEnvAttr(
                       EnvironmentHandle,
                       SQL_ATTR_ODBC_VERSION,
                       reinterpret_cast<SQLPOINTER>( SQL_OV_ODBC3 ),
                       SQL_IS_INTEGER );
    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn,
                      SQL_HANDLE_ENV, EnvironmentHandle, &ErrorMessage );
        SQLFreeHandle( SQL_HANDLE_ENV, EnvironmentHandle );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQL_OV_ODBC3" );
    }

    // Must catch all errors so can free the environment handle
    try
    {
        // Get the first data source, buffer length is in characters
        memset( DriverDescription, 0, sizeof ( DriverDescription ) );
        memset( DriverAttributes,  0, sizeof ( DriverAttributes  ) );
        DescriptionLengthPtr = 0;
        AttributesLengthPtr  = 0;
        sqlReturn = SQLDrivers( EnvironmentHandle,
                                SQL_FETCH_FIRST,
                                DriverDescription,
                                ARRAYSIZE( DriverDescription ),
                                &DescriptionLengthPtr,
                                DriverAttributes,
                                ARRAYSIZE( DriverAttributes ),
                                &AttributesLengthPtr );
        DriverDescription[ ARRAYSIZE( DriverDescription ) - 1 ] = PXS_CHAR_NULL;
        DriverAttributes[ ARRAYSIZE( DriverAttributes ) - 1 ] = PXS_CHAR_NULL;

        // Add to the list and loop for the rest
        while ( SQL_SUCCEEDED( sqlReturn ) )
        {
            // Add the driver name to the array
            if ( wcslen( DriverDescription ) > 0 )
            {
                pDrivers->Add( DriverDescription );
            }

            // Get the next one
            memset( DriverDescription, 0, sizeof ( DriverDescription ) );
            memset( DriverAttributes,  0, sizeof ( DriverAttributes  ) );
            DescriptionLengthPtr = 0;
            AttributesLengthPtr  = 0;
            sqlReturn = SQLDrivers( EnvironmentHandle,
                                    SQL_FETCH_NEXT,
                                    DriverDescription,
                                    ARRAYSIZE( DriverDescription ),
                                    &DescriptionLengthPtr,
                                    DriverAttributes,
                                    ARRAYSIZE( DriverAttributes ),
                                    &AttributesLengthPtr );
            DriverDescription[ ARRAYSIZE( DriverDescription ) - 1 ] = 0x00;
            DriverAttributes[ ARRAYSIZE( DriverAttributes ) - 1 ]   = 0x00;
        }
    }
    catch ( const Exception& )
    {
        SQLFreeHandle( SQL_HANDLE_ENV, EnvironmentHandle );
        throw;
    }
    SQLFreeHandle( SQL_HANDLE_ENV, EnvironmentHandle );
    pDrivers->Sort( true );
}

//===============================================================================================//
//  Description:
//      Get the user or system data source names
//
//  Parameters:
//      wantSystem - true to return system data source names, otherwise
//                   will return user data source names
//      pNameValues- receives the data source names and description pairs
//
//  Returns:
//      void, note there may be no data sources
//===============================================================================================//
void Odbc::GetSystemOrUserDSNs( bool wantSystem,
                                TArray< NameValue >* pNameValues )
{
    HENV         EnvironmentHandle = nullptr;
    String       ErrorMessage, Name, Value;
    SQLWCHAR     ServerName[ SQL_MAX_DSN_LENGTH + 1 ];
    SQLWCHAR     Description[ 256 ];             // Enough for a description
    SQLRETURN    sqlReturn = 0;
    NameValue    Element;
    SQLSMALLINT  NameLength1Ptr = 0, NameLength2Ptr = 0;
    SQLUSMALLINT Direction  = 0;

    if ( pNameValues == nullptr )
    {
        throw ParameterException( L"pNameValues", __FUNCTION__ );
    }
    pNameValues->RemoveAll();

    sqlReturn = SQLAllocHandle( SQL_HANDLE_ENV,
                                nullptr,            // = SQL_NULL_HANDLE
                                &EnvironmentHandle );
    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, L"SQLAllocHandle",  __FUNCTION__ );
    }

    // Set to ODBC v3
    sqlReturn = SQLSetEnvAttr(
                       EnvironmentHandle,
                       SQL_ATTR_ODBC_VERSION,
                       reinterpret_cast<SQLPOINTER>( SQL_OV_ODBC3 ),
                       SQL_IS_INTEGER );
    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn,
                      SQL_HANDLE_ENV, EnvironmentHandle, &ErrorMessage );
        SQLFreeHandle( SQL_HANDLE_ENV, EnvironmentHandle );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED,
                         ErrorMessage.c_str(), "SQL_OV_ODBC3" );
    }

    // Must catch all errors so can free the environment handle
    try
    {
        if ( wantSystem )
        {
            Direction = SQL_FETCH_FIRST_SYSTEM;
        }
        else
        {
            Direction = SQL_FETCH_FIRST_USER;
        }

        do
        {
            memset( ServerName , 0, sizeof ( ServerName  ) );
            memset( Description, 0, sizeof ( Description ) );
            NameLength1Ptr = 0;
            NameLength2Ptr = 0;
            sqlReturn = SQLDataSources( EnvironmentHandle,
                                        Direction,
                                        ServerName,
                                        ARRAYSIZE( ServerName ),
                                        &NameLength1Ptr,
                                        Description,
                                        ARRAYSIZE( Description ),
                                        &NameLength2Ptr );
            if ( SQL_SUCCEEDED( sqlReturn ) && ServerName[ 0 ] )
            {
                ServerName[ ARRAYSIZE( ServerName ) - 1 ]   = PXS_CHAR_NULL;
                Description[ ARRAYSIZE( Description ) - 1 ] = PXS_CHAR_NULL;
                Name  = ServerName;
                Value = Description;
                Element.SetNameValue( Name, Value );
                pNameValues->Add( Element );
            }
            Direction = SQL_FETCH_NEXT;
        } while ( SQL_SUCCEEDED( sqlReturn ) );
    }
    catch ( const Exception& )
    {
        SQLFreeHandle( SQL_HANDLE_ENV, EnvironmentHandle );
        throw;
    }
    SQLFreeHandle( SQL_HANDLE_ENV, EnvironmentHandle );
    PXSSortNameValueArray( pNameValues );
}

//===============================================================================================//
//  Description:
//      Determine if the input is an SQL String type
//
//  Parameters:
//      sqlType - the data type to test
//
//  Returns:
//      void
//===============================================================================================//
bool Odbc::IsSqlStringType( SQLSMALLINT sqlType ) const
{
    if ( ( sqlType == SQL_CHAR         ) ||
         ( sqlType == SQL_VARCHAR      ) ||
         ( sqlType == SQL_LONGVARCHAR  ) ||
         ( sqlType == SQL_WCHAR        ) ||
         ( sqlType == SQL_WVARCHAR     ) ||
         ( sqlType == SQL_WLONGVARCHAR )  )
    {
        return true;
    }
    return false;
}

//===============================================================================================//
//  Description:
//      Make a DSN less connection string
//
//  Parameters:
//      Dbms     - the database management system name
//      Driver   - the ODBC driver name
//      Server   - the value for the SERVER keyword
//      Database - the database name
//      Uid      - the value for the UID keyword
//      Pwd      - the value for the PWD keyword
//      Port     - the port number to connect on
//      Connection  - receives the connection string
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::MakeConnectionString( const String& Dbms,
                                 const String& Driver,
                                 const String& Server,
                                 const String& Database,
                                 const String& Uid,
                                 const String& Pwd,
                                 const String& Port, String& Connection )
{
    Connection = PXS_STRING_EMPTY;
    if ( Dbms.CompareI( PXS_DBMS_NAME_ACCESS ) == 0 )
    {
        MakeAccessConnString( Driver, Database, Uid, Pwd, &Connection );
    }
    else if ( Dbms.CompareI( PXS_DBMS_NAME_MYSQL ) == 0 )
    {
        MakeMySqlConnString( Driver, Server, Database, Port, Uid, Pwd, &Connection );
    }
    else if ( Dbms.CompareI( PXS_DBMS_NAME_SQL_SERVER ) == 0 )
    {
       MakeSqlServerConnString( Driver, Server, Database, Uid, Pwd, &Connection);
    }
    else if ( Dbms.CompareI( PXS_DBMS_NAME_POSTGRE_SQL ) == 0 )
    {
        MakePostGreConnString( Driver, Server, Database, Uid, Pwd, &Connection );
    }
    else
    {
        throw ParameterException( L"Dbms", __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Convert the an ODBC numeric type to a double
//
//  Parameters:
//      pNumeric - numeric structure holding the value to convert
//
//  Returns:
//      the double value
//===============================================================================================//
double Odbc::SqlNumericStructToDouble( const SQL_NUMERIC_STRUCT* pNumeric )
{
    size_t i = 0;
    double value = 0.0, multiplier = 0.0, temp = 0.0;

    if ( pNumeric == nullptr )
    {
        return 0.0;
    }

    multiplier = 1.0;
    for ( i = 0; i < ARRAYSIZE( pNumeric->val ); i++ )
    {
        temp  = PXSMultiplyDouble( pNumeric->val[ i ], multiplier );
        value = PXSAddDouble( value, temp );
        multiplier *= 256.0;  // Maximum is 16^256 = 3.4E38, i.e. no overflow
    }

    // Test for the sign, 0 and 2 are negative for ODBC v3.0 and v3.5
    // respectively
    if ( ( pNumeric->sign == 0 ) || ( pNumeric->sign == 2 ) )
    {
       value = -value;
    }

    return value;
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLAllocHandle
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::AllocHandle( SQLSMALLINT HandleType,
                        SQLHANDLE   InputHandle,
                        SQLHANDLE*  OutputHandlePtr )
{
    String      ErrorMessage;
    SQLRETURN   sqlReturn = 0;
    SQLSMALLINT contextHandleType = 0;

    sqlReturn = SQLAllocHandle( HandleType, InputHandle, OutputHandlePtr );
    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        switch( HandleType )
        {
            default:
                contextHandleType = 0;
            break;

            case SQL_HANDLE_ENV:
                contextHandleType = 0;
            break;

            case SQL_HANDLE_DBC:
                contextHandleType = SQL_HANDLE_ENV;
            break;

            case SQL_HANDLE_STMT:
                contextHandleType = SQL_HANDLE_DBC;
            break;

            case SQL_HANDLE_DESC:
                contextHandleType = SQL_HANDLE_DBC;
            break;
        }
        MakeErrorMsg( sqlReturn,
                      contextHandleType, InputHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLAllocHandle" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLBindCol
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::BindCol( SQLHSTMT     StatementHandle,
                    SQLUSMALLINT ColumnNumber,
                    SQLSMALLINT  TargetType,
                    SQLPOINTER   TargetValue,
                    SQLLEN       BufferLength,
                    SQLLEN*      StrLen_or_Ind )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLBindCol( StatementHandle,
                            ColumnNumber,
                            TargetType,
                            TargetValue, BufferLength, StrLen_or_Ind );

    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLBindCol SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn,
                      SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLBindCol" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLBindParameter
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::BindParameter( SQLHSTMT     hstmt,
                            SQLUSMALLINT ipar,
                            SQLSMALLINT  fParamType,
                            SQLSMALLINT  fCType,
                            SQLSMALLINT  fSqlType,
                            SQLULEN      cbColDef,
                            SQLSMALLINT  ibScale,
                            SQLPOINTER   rgbValue,
                            SQLLEN       cbValueMax,
                            SQLLEN*      pcbValue )
{
    String  ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLBindParameter( hstmt,
                                    ipar,
                                    fParamType,
                                    fCType,
                                    fSqlType,
                                    cbColDef,
                                    ibScale, rgbValue, cbValueMax, pcbValue );
    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, hstmt, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                           PXS_ERROR_DB_OPERATION_FAILED,
                           ErrorMessage.c_str(), "SQLBindParameter" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLBulkOperations
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::BulkOperations( SQLHSTMT StatementHandle, SQLSMALLINT Operation )
{
    String    ErrorMessage, Temp;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLBulkOperations( StatementHandle, Operation );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn,
                      SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLBulkOperations SQL_SUCCESS_WITH_INFO: '%%1'.",
                        ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        ErrorMessage = L"SQLBulkOperations\r\n\r\n";
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &Temp );
        ErrorMessage += Temp;
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLColAttribute
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::ColAttribute( SQLHSTMT     StatementHandle,
                         SQLUSMALLINT ColumnNumber,
                         SQLUSMALLINT FieldIdentifier,
                         SQLPOINTER   CharacterAttribute,
                         SQLSMALLINT  BufferLength,
                         SQLSMALLINT* StringLength,
                         SQLLEN*      NumericAttribute )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLColAttribute( StatementHandle,
                                 ColumnNumber,
                                 FieldIdentifier,
                                 CharacterAttribute,
                                 BufferLength,
                                 StringLength, NumericAttribute );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLColAttribute SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLColAttribute" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLDescribeCol
//
//  Parameters:
//      See the ODBC documentation
//
//  Remarks:
//      The input strings are not necessarily terminated.
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::DescribeCol( SQLHSTMT     StatementHandle,
                        SQLUSMALLINT ColumnNumber,
                        LPWSTR       ColumnName,
                        SQLSMALLINT  BufferLength,
                        SQLSMALLINT* NameLength,
                        SQLSMALLINT* DataType,
                        SQLULEN*     ColumnSize,
                        SQLSMALLINT* DecimalDigits,
                        SQLSMALLINT* Nullable )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLDescribeCol( StatementHandle,
                                ColumnNumber,
                                ColumnName,
                                BufferLength,
                                NameLength,
                                DataType,
                                ColumnSize, DecimalDigits, Nullable );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLDescribeCol SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLDescribeCol" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLDisconnect
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::Disconnect( SQLHDBC ConnectionHandle )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLDisconnect( ConnectionHandle );
    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_DBC, ConnectionHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLDisconnect" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLDriverConnect
//
//  Parameters:
//      See the ODBC documentation
//      pOutConnectionString - receives the output connection string minus
//                             the password value
//
//  Remarks:
//      The input string is not necessarily terminated.
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::DriverConnect( SQLHDBC       ConnectionHandle,
                          SQLHWND       WindowHandle,
                          const String& ConnectionString,
                          String*       pOutConnectionString,
                          SQLUSMALLINT  DriverCompletion )
{
    size_t    i = 0, numChars = 0, numTokens = 0;
    wchar_t   szOutConnection[ 1024 ] = { 0 };  // Even number at least 1024
    wchar_t*  InConnectionString = nullptr;
    String    ErrorMessage, Temp, Token;
    SQLRETURN sqlReturn = 0;
    SQLSMALLINT    StringLength2 = 0;
    StringArray    Tokens;
    AllocateWChars AllocConnection;

    // Must not pass an empty/null string to SQLDriverConnect
    if ( ConnectionString.IsEmpty() )
    {
        throw ParameterException( L"ConnectionString", __FUNCTION__ );
    }

    if ( pOutConnectionString == nullptr )
    {
        throw ParameterException( L"pOutConnectionString", __FUNCTION__ );
    }
    *pOutConnectionString = PXS_STRING_EMPTY;

    // SQLDriverConnect wants a non-const pointer
    numChars = ConnectionString.GetLength();
    numChars = PXSAddSizeT( numChars, 1 );
    InConnectionString = AllocConnection.New( numChars );
    StringCchCopy( InConnectionString, numChars, ConnectionString.c_str() );

    sqlReturn = SQLDriverConnect( ConnectionHandle,
                                  WindowHandle,
                                  InConnectionString,
                                  SQL_NTS,
                                  szOutConnection,
                                  ARRAYSIZE( szOutConnection ),
                                  &StringLength2, DriverCompletion );
    szOutConnection[ ARRAYSIZE( szOutConnection ) - 1 ] = PXS_CHAR_NULL;

    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn,
                      SQL_HANDLE_DBC, ConnectionHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLDriverConnect SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        // SQL_NO_DATA signifies user cancel
        if ( sqlReturn == SQL_NO_DATA )
        {
            throw SystemException( ERROR_CANCELLED, L"SQL_NO_DATA", "SQLDriverConnect" );
        }
        else
        {
            MakeErrorMsg( sqlReturn,
                          SQL_HANDLE_DBC, ConnectionHandle, &ErrorMessage );
            throw Exception( PXS_ERROR_TYPE_APPLICATION,
                             PXS_ERROR_DB_OPERATION_FAILED,
                             ErrorMessage.c_str(), "SQLDriverConnect" );
        }
    }

    // Strip the the password from the output connection string
    Temp = szOutConnection;
    Temp.ToArray( ';', &Tokens );
    numTokens = Tokens.GetSize();
    for ( i = 0; i < numTokens; i++ )
    {
        Token = Tokens.Get( i );
        Token.Trim();
        if ( Token.StartsWith( L"PWD", false ) == false )
        {
            *pOutConnectionString += Token;
            *pOutConnectionString += L";";
        }
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLEndTran
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::EndTran( SQLSMALLINT HandleType,
                    SQLHANDLE   Handle,
                    SQLSMALLINT CompletionType )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLEndTran( HandleType, Handle, CompletionType );
    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, HandleType, Handle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLEndTran" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLExecDirect
//
//  Parameters:
//      See the ODBC documentation
//
//  Remarks:
//      The input string is not necessarily terminated.
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::ExecDirect( SQLHSTMT StatementHandle, const String& SqlQuery )
{
    size_t    numChars = 0;
    wchar_t*  pszSQL = nullptr;
    String    ErrorMessage, Temp;
    SQLRETURN sqlReturn = 0;
    AllocateWChars AllocSQL;

    // Do not pass an empty/null string to SQLExecDirect
    numChars = SqlQuery.GetLength();
    if ( numChars == 0 )
    {
        throw ParameterException( L"SqlQuery", __FUNCTION__ );
    }

    // SQLExecDirect wants a non-const pointer
    numChars = PXSAddSizeT( numChars, 1 );       // + Terminator
    pszSQL   = AllocSQL.New( numChars );
    StringCchCopy( pszSQL, numChars, SqlQuery.c_str() );
    PXSLogAppInfo1( L"Odbc::ExecDirect: '%%1'", SqlQuery );

    // DELETE queries return SQL_NO_DATA if there is nothing to delete
    sqlReturn = SQLExecDirect( StatementHandle, pszSQL, SQL_NTS );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn,
                      SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLExecDirect SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( ( !SQL_SUCCEEDED( sqlReturn ) ) && ( sqlReturn != SQL_NO_DATA ) )
    {
        ErrorMessage = L"SQLExecDirect\r\n\r\n";
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &Temp );
        ErrorMessage += Temp;
        ErrorMessage += PXS_STRING_CRLF;
        ErrorMessage += L"\r\nSQL:\r\n";
        ErrorMessage += SqlQuery;
        ErrorMessage += PXS_STRING_CRLF;
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLExecute
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::Execute( SQLHSTMT StatementHandle )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLExecute( StatementHandle );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLExecute SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLExecute" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLFetch
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      SQL error code
//===============================================================================================//
SQLRETURN Odbc::Fetch( SQLHSTMT StatementHandle )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLFetch( StatementHandle );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLFetch SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( ( !SQL_SUCCEEDED( sqlReturn ) ) && ( sqlReturn != SQL_NO_DATA ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLFetch" );
    }

    return sqlReturn;
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLFetchScroll
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      SQL error code
//===============================================================================================//
SQLRETURN Odbc::FetchScroll( SQLHSTMT    StatementHandle,
                             SQLSMALLINT FetchOrientation,
                             SQLLEN      FetchOffset )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLFetchScroll( StatementHandle, FetchOrientation, FetchOffset );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLFetchScroll SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( ( !SQL_SUCCEEDED( sqlReturn ) ) && ( sqlReturn != SQL_NO_DATA ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLFetchScroll" );
    }

    return sqlReturn;
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLFreeHandle
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::FreeHandle( SQLSMALLINT HandleType, SQLHANDLE Handle )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLFreeHandle( HandleType, Handle );
    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, HandleType, Handle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLFreeHandle" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLFreeStmt
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::FreeStmt( SQLHSTMT StatementHandle, SQLUSMALLINT Option )

{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLFreeStmt( StatementHandle, Option );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLFreeStmt SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLFreeStmt" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLGetConnectAttr
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::GetConnectAttr( SQLHDBC      ConnectionHandle,
                           SQLINTEGER   Attribute,
                           SQLPOINTER   ValuePtr,
                           SQLINTEGER   BufferLength,
                           SQLINTEGER*  StringLengthPtr )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLGetConnectAttr( ConnectionHandle,
                                   Attribute,
                                   ValuePtr,
                                   BufferLength, StringLengthPtr );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_DBC, ConnectionHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLGetConnectAttr SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_DBC, ConnectionHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLGetConnectAttr" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLGetInfo
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::GetInfo( SQLHDBC      ConnectionHandle,
                    SQLUSMALLINT InfoType,
                    SQLPOINTER   InfoValuePtr,
                    SQLSMALLINT  BufferLength,
                    SQLSMALLINT* StringLengthPtr )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLGetInfo( ConnectionHandle,
                            InfoType,
                            InfoValuePtr, BufferLength, StringLengthPtr );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_DBC, ConnectionHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLGetInfo SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_DBC, ConnectionHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLGetInfo" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLGetStmtAttr
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::GetStmtAttr( SQLHSTMT    StatementHandle,
                        SQLINTEGER  Attribute,
                        SQLPOINTER  ValuePtr,
                        SQLINTEGER  BufferLength,
                        SQLINTEGER* StringLengthPtr )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLGetStmtAttr( StatementHandle,
                                Attribute,
                                ValuePtr, BufferLength, StringLengthPtr );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLGetStmtAttr SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLGetStmtAttr" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLGetTypeInfo
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::GetTypeInfo( SQLHSTMT StatementHandle, SQLSMALLINT DataType )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLGetTypeInfo( StatementHandle, DataType );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLGetTypeInfo SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLGetTypeInfo" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLNumResultCols
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::NumResultCols( SQLHSTMT     StatementHandle,
                          SQLSMALLINT* ColumnCountPtr )
{
    String  ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLNumResultCols( StatementHandle, ColumnCountPtr );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLNumResultCols SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLNumResultCols" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLPrepare
//
//  Parameters:
//      See the ODBC documentation
//
//  Remarks:
//      Input string is not necessarily terminated
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::Prepare( SQLHSTMT    StatementHandle,
                    SQLWCHAR*   StatementText,
                    SQLINTEGER  TextLength )
{
    size_t     lenChars = 0;
    String     ErrorMessage;
    SQLWCHAR   szEmpty[ 1 ] = { PXS_CHAR_NULL };  // i.e. "";
    SQLWCHAR*  pStatement = StatementText;
    SQLRETURN  sqlReturn = 0;
    SQLINTEGER length    = TextLength;

    // Should not pass to SQLPrepare a NULL for StatementText
    if ( pStatement == nullptr )
    {
        pStatement = szEmpty;  // i.e. "";
        length     = 0;
    }

    sqlReturn = SQLPrepare( StatementHandle, pStatement, length );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLPrepare SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        if ( StatementText && ( TextLength > 0 ) )
        {
            lenChars = PXSCastInt64ToSizeT( TextLength );
            ErrorMessage.AppendChars( StatementText, lenChars );
        }
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLPrepare" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLProcedures
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//
//===============================================================================================//
void Odbc::Procedures( SQLHSTMT StatementHandle,
                       const String& CatalogName,
                       const String& SchemaName,
                       const String& ProcName )
{
    size_t    numChars    = 0;
    wchar_t*  pszCatalogName = nullptr;
    wchar_t*  pszSchemaName  = nullptr;
    wchar_t*  pszProcName    = nullptr;
    String    ErrorMessage;
    SQLRETURN sqlReturn   = 0;
    SQLSMALLINT    NameLength1 = 0, NameLength2 = 0, NameLength3 = 0;
    AllocateWChars AllocCatalog, AllocSchema, AllocProc;

    // SQLProcedures wants non-const strings
    numChars = CatalogName.GetLength();
    if ( numChars )
    {
        NameLength1    = SQL_NTS;
        numChars       = PXSAddSizeT( numChars, 1 );
        pszCatalogName = AllocCatalog.New( numChars );
        StringCchCopy( pszCatalogName, numChars, CatalogName.c_str() );
    }

    numChars = SchemaName.GetLength();
    if ( numChars )
    {
        NameLength2   = SQL_NTS;
        numChars      = PXSAddSizeT( numChars, 1 );
        pszSchemaName = AllocSchema.New( numChars );
        StringCchCopy( pszSchemaName, numChars, SchemaName.c_str() );
    }

    numChars = ProcName.GetLength();
    if ( numChars )
    {
        NameLength3 = SQL_NTS;
        numChars    = PXSAddSizeT( numChars, 1 );
        pszProcName = AllocProc.New( numChars );
        StringCchCopy( pszProcName, numChars, ProcName.c_str() );
    }

    sqlReturn = SQLProcedures( StatementHandle,
                               pszCatalogName,
                               NameLength1,
                               pszSchemaName,
                               NameLength2, pszProcName, NameLength3 );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLProcedures SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLProcedures" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLRowCount
//
//  Parameters:
//      See the ODBC documentation
//
//  Remarks:
//      SQLRowCount returns signed 32-bit value and signed 64-bit value
//      on 32-bit and 64-bit platforms respectively, i.e. ssize_t
//
//      The second parameter retains the pre-64 bit name so as not to
//      shadow the method name.
//
//  Returns:
//      void
//
//===============================================================================================//
void Odbc::RowCount( SQLHSTMT StatementHandle, SQLLEN* RowCountPtr )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLRowCount( StatementHandle, RowCountPtr );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLRowCount SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLRowCount" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLSetConnectAttr
//
//  Parameters:
//      See the ODBC documentation
//
//  Remarks:
//
//  Returns:
//      void
//
//===============================================================================================//
void Odbc::SetConnectAttr( SQLHDBC    ConnectionHandle,
                           SQLINTEGER Attribute,
                           SQLPOINTER ValuePtr,
                           SQLINTEGER StringLength )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLSetConnectAttr( ConnectionHandle,
                                   Attribute, ValuePtr, StringLength );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_DBC, ConnectionHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SetConnectAttr SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_DBC, ConnectionHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED,
                         ErrorMessage.c_str(), "SQLSetConnectAttr" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLSetDescField
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::SetDescField( SQLHDESC    DescriptorHandle,
                         SQLSMALLINT RecNumber,
                         SQLSMALLINT FieldIdentifier,
                         SQLPOINTER  ValuePtr,
                         SQLINTEGER  BufferLength )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLSetDescField( DescriptorHandle,
                                 RecNumber,
                                 FieldIdentifier, ValuePtr, BufferLength );
    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_DESC, DescriptorHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLSetDescField" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLSetEnvAttr
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::SetEnvAttr( SQLHENV    EnvironmentHandle,
                       SQLINTEGER Attribute,
                       SQLPOINTER ValuePtr,
                       SQLINTEGER StringLength )
{
    String    ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLSetEnvAttr( EnvironmentHandle,
                               Attribute, ValuePtr, StringLength );
    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_ENV, EnvironmentHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLSetEnvAttr" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLSetStmtAttr
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::SetStmtAttr( SQLHSTMT   StatementHandle,
                        SQLINTEGER Attribute,
                        SQLPOINTER ValuePtr,
                        SQLINTEGER StringLength )
{
    String  ErrorMessage;
    SQLRETURN sqlReturn = 0;

    sqlReturn = SQLSetStmtAttr( StatementHandle,
                                Attribute, ValuePtr, StringLength );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLSetStmtAttr SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLSetStmtAttr" );
    }
}

//===============================================================================================//
//  Description:
//      Wrapper for SQLTables
//
//  Parameters:
//      See the ODBC documentation
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::Tables( SQLHSTMT  StatementHandle,
                   const String& CatalogName,
                   const String& SchemaName,
                   const String& TableName,
                   const String& TableType )
{
    size_t   numChars = 0;
    wchar_t* pszCatalogName = nullptr;
    wchar_t* pszSchemaName  = nullptr;
    wchar_t* pszTableName   = nullptr;
    wchar_t* pszTableType   = nullptr;
    String      ErrorMessage;
    SQLRETURN   sqlReturn   = 0;
    SQLSMALLINT NameLength1 = 0, NameLength2 = 0;
    SQLSMALLINT NameLength3 = 0, NameLength4 = 0;
    AllocateWChars AllocCatalog, AllocSchema, AllocTable, AllocType;

    // SQLTables wants non-const strings
    numChars = CatalogName.GetLength();
    if ( numChars )
    {
        NameLength1    = SQL_NTS;
        numChars       = PXSAddSizeT( numChars, 1 );        // Terminator
        pszCatalogName = AllocCatalog.New( numChars );
        StringCchCopy( pszCatalogName, numChars, CatalogName.c_str() );
    }

    numChars = SchemaName.GetLength();
    if ( numChars )
    {
        NameLength2   = SQL_NTS;
        numChars      = PXSAddSizeT( numChars, 1 );
        pszSchemaName = AllocSchema.New( numChars );
        StringCchCopy( pszSchemaName, numChars, SchemaName.c_str() );
    }

    numChars = TableName.GetLength();
    if ( numChars )
    {
        NameLength3  = SQL_NTS;
        numChars     = PXSAddSizeT( numChars, 1 );
        pszTableName = AllocTable.New( numChars );
        StringCchCopy( pszTableName, numChars, TableName.c_str() );
    }

    numChars = TableType.GetLength();
    if ( numChars )
    {
        NameLength4  = SQL_NTS;
        numChars     = PXSAddSizeT( numChars, 1 );
        pszTableType = AllocType.New( numChars );
        StringCchCopy( pszTableType, numChars, TableType.c_str() );
    }

    sqlReturn = SQLTables( StatementHandle,
                           pszCatalogName ,
                           NameLength1,
                           pszSchemaName,
                           NameLength2,
                           pszTableName,
                           NameLength3, pszTableType, NameLength4 );
    if ( sqlReturn == SQL_SUCCESS_WITH_INFO )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        PXSLogAppInfo1( L"SQLTables SQL_SUCCESS_WITH_INFO: '%%1'.", ErrorMessage );
    }

    if ( !SQL_SUCCEEDED( sqlReturn ) )
    {
        MakeErrorMsg( sqlReturn, SQL_HANDLE_STMT, StatementHandle, &ErrorMessage );
        throw Exception( PXS_ERROR_TYPE_APPLICATION,
                         PXS_ERROR_DB_OPERATION_FAILED, ErrorMessage.c_str(), "SQLTables" );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Make a connection string for Microsoft Access
//
//  Parameters:
//      Driver      - the ODBC driver name, if none supplied will find one
//      Database    - the database name
//      Uid         - the user id
//      Pwd         - the user's password
//      pConnection - receives the string;
//
//  Remarks:
//      On a clean install of Windows 10, there are no problems connecting
//      to an access database. However, after installing/updating Office 365
//      the ODBC driver manager says cannot load odbcji32.dll and there
//      is an assocated SQL_HANDLE_ENV error in SQLAllocHandle. Can connect
//      if use "DRIVER=Microsoft Access Driver (*.mdb)"
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::MakeAccessConnString( const String& Driver,
                                 const String& Database,
                                 const String& Uid,
                                 const String& Pwd, String* pConnection )
{
    String      OdbcDriver;
    Formatter   Format;
    StringArray Drivers;
    SystemInformation SysInfo;

    if ( pConnection == nullptr )
    {
        throw ParameterException( L"pConnection", __FUNCTION__ );
    }
    *pConnection = PXS_STRING_EMPTY;

    // DRIVER, if have driver use it otherwise find an installed one
    OdbcDriver = Driver;
    OdbcDriver.Trim();
    if ( OdbcDriver.IsEmpty() )
    {
        // To v2007  : Microsoft Access Driver (*.mdb)
        // From v2007: Microsoft Access Driver (*.mdb, *.accdb)
        if ( Database.EndsWithStringI( L".accdb" ) )
        {
            OdbcDriver = L"Microsoft Access Driver (*.mdb, *.accdb)";
        }
        else
        {
            // For mdb databases, use the newer (*.mdb, *.accdb) driver if
            // present but not on Windows 10
            OdbcDriver = L"Microsoft Access Driver (*.mdb)";
            GetInstalledDrivers( &Drivers );
            if ( ( SysInfo.GetMajorVersion() < 10 ) &&
                 ( Drivers.IndexOf( L"Microsoft Access Driver (*.mdb, *.accdb)",
                                                    false ) != PXS_MINUS_ONE ) )
            {
                OdbcDriver = L"Microsoft Access Driver (*.mdb, *.accdb)";
            }
        }
    }
    *pConnection += Format.String1( L"DRIVER={%%1};", OdbcDriver );

    // DBQ
    if ( Database.GetLength() )
    {
        *pConnection += Format.String1( L"DBQ=%%1;", Database );
    }

    // UID, only add user id if have it otherwise the connection string can
    // fail if a default user id is used, e.g. Admin for Access.
    if ( Uid.GetLength() )
    {
        *pConnection += Format.String1( L"UID=%%1;", Uid );
    }

    // PWD
    if ( Pwd.GetLength() )
    {
        *pConnection += Format.String1( L"PWD=%%1;", Pwd );
    }
}

//===============================================================================================//
//  Description:
//      Make an SQL error message
//
//  Parameters:
//      sqlReturn    - the sql error code of the ODBC function
//      HandleType   - handle type
//      Handle       - SQL handle
//      ErrorMessage - receives the error message
//
//  Remarks:
//      Call SQLGetDiagRec repeatedly until no more data. SQLGetDiagRec can
//      take an arbitrarily large buffer but just use SQL_MAX_MESSAGE_LENGTH
//      else will need to loop for each nRecNumber to concat to form a
//      single message.
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::MakeErrorMsg( SQLRETURN   sqlReturn,
                         SQLSMALLINT HandleType,
                         SQLHANDLE   Handle, String* pErrorMessage )
{
    const SQLSMALLINT MAX_RECORDS = 10;  // Get up to 10 records
    size_t      i = 0;
    String      Text;
    SQLWCHAR    SQLState[ SQL_SQLSTATE_SIZE + 1 ] = { 0 };
    SQLWCHAR    MessageText[ SQL_MAX_MESSAGE_LENGTH + 1 ] = { 0 };
    Formatter   Format;
    SQLRETURN   result = 0;
    SQLINTEGER  NativeError = 0;
    SQLSMALLINT TextLength  = 0, RecNumber = 0;

    if ( pErrorMessage == nullptr )
    {
        return;     // Cannot do anything
    }
    pErrorMessage->Allocate( 256 );
    *pErrorMessage = PXS_STRING_EMPTY;

    // Get the first message
    RecNumber = 1;      // Counting starts at 1
    result = SQLGetDiagRec( HandleType,
                            Handle,
                            RecNumber,
                            SQLState,
                            &NativeError,
                            MessageText,
                            ARRAYSIZE( MessageText ), &TextLength );
    SQLState[ ARRAYSIZE( SQLState ) - 1 ] = PXS_CHAR_NULL;
    MessageText[ ARRAYSIZE( MessageText ) - 1 ] = PXS_CHAR_NULL;

    // Build the cumulative error string
    while ( ( result    != SQL_NO_DATA ) &&            // = 100
            ( result    != SQL_ERROR   ) &&            // = -1
            ( result    != SQL_INVALID_HANDLE ) &&     // = -2
            ( RecNumber <= MAX_RECORDS ) )
    {
        // See if this is a continued error message
        if ( RecNumber > 1  )
        {
            *pErrorMessage += L"\r\n\r\n====================\r\n";
        }
        *pErrorMessage += MessageText;        // SQLCHAR == unsigned char
        *pErrorMessage += PXS_STRING_CRLF;

        PXSGetResourceString( PXS_IDS_122_SQL_STATE, &Text );
        *pErrorMessage += Text;
        *pErrorMessage += PXS_CHAR_COLON;
        *pErrorMessage += PXS_CHAR_SPACE;
        *pErrorMessage += SQLState;         // SQLCHAR == unsigned char
        *pErrorMessage += PXS_STRING_CRLF;

        PXSGetResourceString( PXS_IDS_123_NATIVE_ERROR, &Text );
        *pErrorMessage += Text;
        *pErrorMessage += PXS_CHAR_COLON;
        *pErrorMessage += PXS_CHAR_SPACE;
        *pErrorMessage += Format.Int32( NativeError );
        *pErrorMessage += PXS_STRING_CRLF;

        // Increment the record counter and  call again
        RecNumber = PXSAddInt16( RecNumber, 1 );
        memset( SQLState, 0, sizeof ( SQLState ) );
        memset( MessageText, 0, sizeof ( MessageText ) );
        result = SQLGetDiagRec( HandleType,
                                Handle,
                                RecNumber,
                                SQLState,
                                &NativeError,
                                MessageText,
                                ARRAYSIZE( MessageText ), &TextLength );

        SQLState[ ARRAYSIZE( SQLState ) - 1 ] = PXS_CHAR_NULL;
        MessageText[ ARRAYSIZE( MessageText ) - 1 ] = PXS_CHAR_NULL;
    }

    // Add the sql return error code
    PXSGetResourceString( PXS_IDS_100_ERROR_NUMBER, &Text );
    *pErrorMessage += PXS_STRING_CRLF;
    *pErrorMessage += L"________________________________\r\n\r\n";
    *pErrorMessage += Text;
    *pErrorMessage += PXS_CHAR_COLON;
    *pErrorMessage += PXS_CHAR_SPACE;
    *pErrorMessage += Format.Int32( sqlReturn );  // SQLRETURN = SQLSMALLINT

    struct _ERROR_NAMES
    {
        SQLSMALLINT sqlReturn;
        LPCWSTR     pszName;
    } Names[] =
          { { SQL_SUCCESS           , L" [SQL_SUCCESS]"           } ,
            { SQL_SUCCESS_WITH_INFO , L" [SQL_SUCCESS_WITH_INFO]" } ,
            { SQL_NEED_DATA         , L" [SQL_NEED_DATA]"         } ,
            { SQL_STILL_EXECUTING   , L" [SQL_STILL_EXECUTING]"   } ,
            { SQL_ERROR             , L" [SQL_ERROR]"             } ,
            { SQL_INVALID_HANDLE    , L" [SQL_INVALID_HANDLE]"    } };

    for ( i = 0; i < ARRAYSIZE( Names ); i++ )
    {
        if ( sqlReturn == Names[ i ].sqlReturn )
        {
            *pErrorMessage += Names[ i ].pszName;
            break;
        }
    }
    *pErrorMessage += PXS_STRING_DOT;
    pErrorMessage->Trim();
}

//===============================================================================================//
//  Description:
//      Make a connection string for MySQL
//
//  Parameters:
//      Driver      - the ODBC driver name
//      Server      - the server name
//      Database    - the database name
//      Port        - the port number
//      Uid         - the user id
//      Pwd         - the user's password
//      pConnection - receives the string;
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::MakeMySqlConnString( const String& Driver,
                                const String& Server,
                                const String& Database,
                                const String& Port,
                                const String& Uid,
                                const String& Pwd, String* pConnection )
{
    String    OdbcDriver;
    Formatter Format;

    if ( pConnection == nullptr )
    {
        throw ParameterException( L"pConnection", __FUNCTION__ );
    }
    *pConnection = PXS_STRING_EMPTY;

    // DRIVER, the newer v5 seems to have problem showing the login dialogue
    // when using SQLDriverConnect so if non specified use v3.51
    OdbcDriver = Driver;
    OdbcDriver.Trim();
    if ( OdbcDriver.IsEmpty() )
    {
        OdbcDriver = L"MySQL ODBC 3.51 Driver";
    }
    *pConnection += Format.String1( L"DRIVER={%%1};", OdbcDriver );

    // SERVER
    if ( Server.GetLength() )
    {
        *pConnection += Format.String1( L"SERVER=%%1;", Server );
    }

    // DATABASE
    if ( Database.GetLength() )
    {
        *pConnection += Format.String1( L"DATABASE=%%1;", Database );
    }

    // PORT
    if ( Port.GetLength() )
    {
        *pConnection += Format.String1( L"PORT=%%1;", Port );
    }

    // OPTION, as recommended will use the same as for Visual basic = 3.
    // FLAG_FIELD_LENGTH (1) prevents MySQL from reporting the real column
    // width. FLAG_FOUND_ROWS (2) which forces MySQL to return the matching
    // rows. Tell the driver not to cache the result (1048576), otherwise
    // the server will send the entire result set rather than when call fetch.
    // 1 + 2 + 1048576 = 1048579
    *pConnection += L"OPTION=1048579;";

    // UID
    if ( Uid.GetLength() )
    {
        *pConnection += Format.String1( L"UID=%%1;", Uid );
    }

    // PWD
    if ( Pwd.GetLength() )
    {
        *pConnection += Format.String1( L"PWD=%%1;", Pwd );
    }
}

//===============================================================================================//
//  Description:
//      Make a connection string for Postgre SQL
//
//  Parameters:
//      Driver      - the ODBC driver name
//      Server      - the server name
//      Database    - the database name
//      Uid         - the user id
//      Pwd         - the user's password
//      pConnection - receives the string;
//
//  Remarks:
//      SERVER, will not use as when it is part of the connection string the
//      text field on the login prompt is disabled preventing the user from
//      selecting a different server.
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::MakePostGreConnString( const String& Driver,
                                  const String& /* Server */,
                                  const String& Database,
                                  const String& Uid, const String& Pwd, String* pConnection )
{
    Formatter Format;

    if ( pConnection == nullptr )
    {
        throw ParameterException( L"pConnection", __FUNCTION__ );
    }
    *pConnection = PXS_STRING_EMPTY;

    // DRIVER, must have this;
    if ( Driver.IsEmpty() )
    {
        throw ParameterException( L"Driver", __FUNCTION__ );
    }
    *pConnection += Format.String1( L"DRIVER={%%1};", Driver );

    // DATABASE
    if ( Database.GetLength() )
    {
      *pConnection += Format.String1( L"DATABASE=%%1;", Database);
    }

    // UID
    if ( Uid.GetLength() )
    {
        *pConnection += Format.String1( L"UID=%%1;", Uid );
    }

    // PWD, SQL Server seems to always want this even if none supplied
    if ( Pwd.GetLength() )
    {
        *pConnection += Format.String1( L"PWD=%%1;", Pwd );
    }
}

//===============================================================================================//
//  Description:
//      Make a connection string for SQL Server
//
//  Parameters:
//      Driver      - the ODBC driver name
//      Server      - the server name
//      Database    - the database name
//      Uid         - the user id
//      Pwd         - the user's password
//      pConnection - receives the string;
//
//  Returns:
//      void
//===============================================================================================//
void Odbc::MakeSqlServerConnString( const String& Driver,
                                    const String& Server,
                                    const String& Database,
                                    const String& Uid, const String& Pwd, String* pConnection )
{
    String    OdbcDriver, PwdClean;
    Formatter Format;

    if ( pConnection == nullptr )
    {
        throw ParameterException( L"pConnection", __FUNCTION__ );
    }
    *pConnection = PXS_STRING_EMPTY;

    // DRIVER, if none supplied, use the expected value
    OdbcDriver = Driver;
    OdbcDriver.Trim();
    if ( OdbcDriver.IsEmpty() )
    {
        OdbcDriver = L"SQL Server";
    }
    *pConnection += Format.String1( L"DRIVER={%%1};", OdbcDriver );

    // SERVER
    if ( Server.GetLength() )
    {
        *pConnection += Format.String1( L"SERVER=%%1;", Server );
    }

    // DATABASE
    if ( Database.GetLength() )
    {
      *pConnection += Format.String1( L"DATABASE=%%1;", Database);
    }

    // UID
    if ( Uid.GetLength() )
    {
        *pConnection += Format.String1( L"UID=%%1;", Uid );
    }

    // PWD, SQL Server seems to always want this
    PwdClean = Pwd;
    if ( PwdClean.IsEmpty() )
    {
        PwdClean = PXS_STRING_EMPTY;
    }
    *pConnection += Format.String1( L"PWD=%%1;", PwdClean );
}
