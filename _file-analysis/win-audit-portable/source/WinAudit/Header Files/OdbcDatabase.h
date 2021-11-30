///////////////////////////////////////////////////////////////////////////////////////////////////
//
// ODBC Database and Wrapper Class Header
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

#ifndef WINAUDIT_ODBC_DATABASE_H_
#define WINAUDIT_ODBC_DATABASE_H_

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
#include "PxsBase/Header Files/StringT.h"

// 5. This Project
#include "WinAudit/Header Files/Odbc.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class OdbcDatabase
{
    public:
        // Default constructor
        OdbcDatabase();

        // Destructor
        ~OdbcDatabase();

        // Methods
        void    BeginTrans();
        void    CommitTrans();
        void    Connect( const String& ConnectionString,
                         DWORD connectTimeoutSecs, DWORD queryTimeoutSecs, HWND hWnd );
        void    Disconnect();
        SQLLEN  ExecuteDirect( const String& SqlQuery );
        bool    ExecuteSelectSqlInteger( const String& SqlQuery, SQLINTEGER* pSqlInteger );
        void    ExecuteTrans( const String& SqlQuery,
                              SQLLEN maxRecordsAffected = PXS_SQLLEN_MAX );
        void    FixUpStringSQL( const String& Input, SQLSMALLINT sqlType, String* pOutput ) const;
        SQLHDBC GetConnectionHandle() const;
        void    GetDbKeyWord( LPCWSTR pszAnsiKeyword, String* pDbKeyword );
        void    GetDatabaseName( String* pDatabaseName );
const String&   GetDbmsName() const;
        void    GetMajorMinorVersion( int* pMajor, int* pMinor ) const;
        void    GetProcedures( const String& Filter, StringArray* pProcedures );
const String&   GetServerName() const;
        bool    GetSupportsSqlType( short sqlType, String* pTypeName );
        void    GetTables( bool wantUser, bool wantSystem, bool wantViews, StringArray* pTables );
const String&   GetUserName() const;
        bool    IsCaseSensitiveSort() const;
        bool    IsConnected() const;
        void    QueryTimeOut( DWORD queryTimeOut );
        DWORD   QueryTimeOut() const;
        void    RollbackTrans();
        bool    StartedTrans();
        bool    SupportsQueryTimeOut() const;
        bool    SupportsStoredProcedures();
        bool    SupportsViews();

    protected:
        // Data members
        HENV        m_EnvironmentHandle;
        SQLHDBC     m_ConnectionHandle;
        SQLHSTMT    m_hTransactionStatement;   // For use within transactions
        Odbc        m_ODBC;

        // Methods

    private:
        // Copy constructor - not allowed
        OdbcDatabase( const OdbcDatabase& oOdbcDatabase );

        // Assignment operator - not allowed
        OdbcDatabase& operator= ( const OdbcDatabase& oOdbcDatabase );

        // Methods
        void    FreeTransactionStatement();
        void    GetKeywordCurrentTimestamp( String* pDbKeyWord );
        void    GetKeywordUpper( String* pDbKeyWord );
        void    GetKeywordUtcTimestamp( String* pDbKeyWord );
        bool    IsDateDataType( SQLSMALLINT sqlType ) const;
        bool    IsStringDataType( SQLSMALLINT sqlType ) const;
        void    SetQueryTimeoutAttr( SQLHSTMT hStmt );

        // Data members
        bool    m_bConnected;
        DWORD   m_uQueryTimeOut;
        String  m_DbmsName;
        String  m_DbmsVersion;
        String  m_DmVersion;
        String  m_DriverName;
        String  m_DriverVersion;
        String  m_OdbcVersion;
        String  m_DataSourceName;
        String  m_ServerName;
        String  m_UserName;
};

#endif  // WINAUDIT_ODBC_DATABASE_H_
