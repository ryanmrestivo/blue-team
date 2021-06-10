///////////////////////////////////////////////////////////////////////////////////////////////////
//
// ODBC Utility and Wrapper Class Header
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

#ifndef WINAUDIT_ODBC_H_
#define WINAUDIT_ODBC_H_

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
class NameValue;
template< class T > class TArray;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class Odbc
{
    public:
        // Default constructor
        Odbc();

        // Destructor
        ~Odbc();

        // Methods
        void    GetFileDataSourceNames( StringArray* pDSNs );
        void    GetInstalledDrivers( StringArray* pDrivers );
        void    GetSystemOrUserDSNs( bool wantSystem,
                                     TArray< NameValue >* pNameValues );
        bool    IsSqlStringType( SQLSMALLINT sqlType ) const;
        void    MakeConnectionString( const String& Dbms,
                                      const String& Driver,
                                      const String& Server,
                                      const String& Database,
                                      const String& Uid,
                                      const String& Pwd,
                                      const String& Port, String& Connection);
        double  SqlNumericStructToDouble( const SQL_NUMERIC_STRUCT* pNumeric );

        // SQL functions
        void AllocHandle( SQLSMALLINT HandleType,
                          SQLHANDLE   InputHandle,
                          SQLHANDLE*  OutputHandlePtr );
        void BindCol( SQLHSTMT     StatementHandle,
                      SQLUSMALLINT ColumnNumber,
                      SQLSMALLINT  TargetType,
                      SQLPOINTER   TargetValue,
                      SQLLEN       BufferLength,
                      SQLLEN*      StrLen_or_Ind );
        void BindParameter( SQLHSTMT     hstmt,
                            SQLUSMALLINT ipar,
                            SQLSMALLINT  fParamType,
                            SQLSMALLINT  fCType,
                            SQLSMALLINT  fSqlType,
                            SQLULEN      cbColDef,
                            SQLSMALLINT  ibScale,
                            SQLPOINTER   rgbValue,
                            SQLLEN       cbValueMax,
                            SQLLEN*      pcbValue );
        void BulkOperations( SQLHSTMT    StatementHandle,
                             SQLSMALLINT Operation );
        void ColAttribute( SQLHSTMT     StatementHandle,
                           SQLUSMALLINT ColumnNumber,
                           SQLUSMALLINT FieldIdentifier,
                           SQLPOINTER   CharacterAttribute,
                           SQLSMALLINT  BufferLength,
                           SQLSMALLINT* StringLength,
                           SQLLEN*      NumericAttribute );
        void DescribeCol( SQLHSTMT     StatementHandle,
                          SQLUSMALLINT ColumnNumber,
                          LPWSTR       ColumnName,
                          SQLSMALLINT  BufferLength,
                          SQLSMALLINT* NameLength,
                          SQLSMALLINT* DataType,
                          SQLULEN*     ColumnSize,
                          SQLSMALLINT* DecimalDigits,
                          SQLSMALLINT* Nullable );
        void Disconnect( SQLHDBC ConnectionHandle );
        void DriverConnect( SQLHDBC       ConnectionHandle,
                            SQLHWND       WindowHandle,
                            const String& ConnectionString,
                            String*       pOutConnectionString,
                            SQLUSMALLINT DriverCompletion );
        void EndTran( SQLSMALLINT HandleType,
                      SQLHANDLE   Handle,
                      SQLSMALLINT CompletionType );
        void ExecDirect( SQLHSTMT StatementHandle, const String& SqlQuery );
        void Execute( SQLHSTMT StatementHandle );
        SQLRETURN Fetch( SQLHSTMT StatementHandle );
        SQLRETURN FetchScroll( SQLHSTMT    StatementHandle,
                               SQLSMALLINT FetchOrientation,
                               SQLLEN      FetchOffset );
        void FreeHandle( SQLSMALLINT HandleType, SQLHANDLE Handle );
        void FreeStmt( SQLHSTMT StatementHandle, SQLUSMALLINT Option );
        void GetConnectAttr( SQLHDBC     ConnectionHandle,
                             SQLINTEGER  Attribute,
                             SQLPOINTER  ValuePtr,
                             SQLINTEGER  BufferLength,
                             SQLINTEGER* StringLengthPtr );
        void GetInfo( SQLHDBC      ConnectionHandle,
                      SQLUSMALLINT InfoType,
                      SQLPOINTER   InfoValuePtr,
                      SQLSMALLINT  BufferLength,
                      SQLSMALLINT* StringLengthPtr );
        void GetStmtAttr( SQLHSTMT    StatementHandle,
                          SQLINTEGER  Attribute,
                          SQLPOINTER  ValuePtr,
                          SQLINTEGER  BufferLength,
                          SQLINTEGER* StringLengthPtr );
        void GetTypeInfo( SQLHSTMT    StatementHandle,
                          SQLSMALLINT DataType);
        void NumResultCols( SQLHSTMT StatementHandle,
                            SQLSMALLINT* ColumnCountPtr);
        void Prepare( SQLHSTMT   StatementHandle,
                      SQLWCHAR*  StatementText,
                      SQLINTEGER TextLength );
        void Procedures( SQLHSTMT StatementHandle,
                         const String& CatalogName,
                         const String& SchemaName,
                         const String& ProcName );
        void RowCount( SQLHSTMT StatementHandle,
                       SQLLEN*  RowCountPtr );
        void SetConnectAttr( SQLHDBC    ConnectionHandle,
                             SQLINTEGER Attribute,
                             SQLPOINTER ValuePtr,
                             SQLINTEGER StringLength );
        void SetDescField( SQLHDESC    DescriptorHandle,
                           SQLSMALLINT RecNumber,
                           SQLSMALLINT FieldIdentifier,
                           SQLPOINTER  ValuePtr,
                           SQLINTEGER  BufferLength );
        void SetEnvAttr( SQLHENV    EnvironmentHandle,
                         SQLINTEGER Attribute,
                         SQLPOINTER ValuePtr,
                         SQLINTEGER StringLength );
        void SetStmtAttr( SQLHSTMT   StatementHandle,
                          SQLINTEGER Attribute,
                          SQLPOINTER ValuePtr,
                          SQLINTEGER StringLength );
        void Tables( SQLHSTMT  StatementHandle,
                     const String& CatalogName,
                     const String& SchemaName,
                     const String& TableName,
                     const String& TableType );
    protected:
        // Methods

        // Data members

    private:
        // Copy constructor - not allowed
        Odbc( const Odbc& oOdbc );

        // Assignment operator - not allowed
        Odbc& operator= ( const Odbc& oOdbc );

        // Methods
        void    MakeAccessConnString( const String& Driver,
                                      const String& Database,
                                      const String& Uid,
                                      const String& Pwd,
                                      String* pConnection );
        void    MakeErrorMsg( SQLRETURN   sqlReturn,
                              SQLSMALLINT HandleType,
                              SQLHANDLE   Handle,
                              String*     pErrorMessage );
        void    MakeMySqlConnString( const String& zDriver,
                                     const String& Server,
                                     const String& Database,
                                     const String& Port,
                                     const String& Uid,
                                     const String& Pwd,
                                     String* pConnection );
        void    MakePostGreConnString( const String& Driver,
                                       const String& Server,
                                       const String& Database,
                                       const String& Uid,
                                       const String& Pwd,
                                       String* pConnection );
        void    MakeSqlServerConnString( const String& Driver,
                                         const String& zServer,
                                         const String& Database,
                                         const String& Uid,
                                         const String& Pwd,
                                         String* pConnection );
       // Data members
};

#endif  // WINAUDIT_ODBC_H_
