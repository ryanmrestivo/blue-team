///////////////////////////////////////////////////////////////////////////////////////////////////
//
// ODBC Record Set Class Header
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

#ifndef WINAUDIT_ODBC_RECORD_SET_H_
#define WINAUDIT_ODBC_RECORD_SET_H_

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
class OdbcDatabase;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class OdbcRecordSet
{
    public:
        // Default constructor
        OdbcRecordSet();

        // Destructor
        ~OdbcRecordSet();

        // Methods
        void        Close();
      const String& FieldValue( size_t idxColumn );
      const String& FieldValue( size_t idxRow, size_t idxColumn );
        size_t      GetColumnCount() const;
        LPCWSTR     GetColumnDisplayName( size_t idxColumn ) const;
        size_t      GetColumnDisplaySizeChars( size_t idxColumn ) const;
      const String& GetColumnName( size_t idxColumn, bool tableDotColumn );
        SQLCHAR     GetColumnPrecision( size_t idxColumn ) const;
        size_t      GetColumnSizeBytes( size_t idxColumn ) const;
        SQLSCHAR    GetColumnScale( size_t idxColumn ) const;
        SQLSMALLINT GetColumnSqlType( size_t idxColumn ) const;
        bool        IsDeleted( size_t idxRow );
        bool        IsFieldNull( size_t idxRow );
        bool        IsFieldNull( size_t idxRow, size_t idxColumn );
        bool        IsOpen() const;
        bool        Move( size_t idxRow );
        bool        MoveNext();
        void        Open( const String& SqlQuery, OdbcDatabase* pDatabase );

    protected:
        // Methods

        // Data members

    private:
        // Structure to hold the column properties. Access has a
        // maximum of 64 characters for table and column names.
        typedef struct _TYPE_COLUMN_PROPERTIES
        {
            bool        boolean;
            SQLSCHAR    scale;                  // char
            SQLCHAR     bPrecision;             // unsigned char
            SQLSMALLINT sqlType;                // short
            SQLSMALLINT cType;                  // short
            size_t      sizeBytes;
            size_t      bytesAllocated;
            size_t      displayChars;
            size_t      rowOffsetBytes;
            size_t      indRowOffsetBytes;
            wchar_t     szName[ 68 ];         // Access = 64, Firebird = 31
            wchar_t     szTable[ 68 ];        // Access = 64, Firebird = 31
            wchar_t     szDisplayName[ 68 ];  // Name used for display purposes
        } TYPE_COLUMN_PROPERTIES;

        // Copy constructor - not allowed
        OdbcRecordSet( const OdbcRecordSet& oOdbcRecordSet );

        // Assignment operator - not allowed
        OdbcRecordSet& operator= ( const OdbcRecordSet& oOdbcRecordSet );

        // Methods
        void    AllocateCache();
        void    BindCache();
        void    DataBufferToString( BYTE* pBuffer, TYPE_COLUMN_PROPERTIES* pColumn );
        void    DeleteColumns();
        size_t  FillCache( SQLSMALLINT FetchOrientation, SQLLEN FetchOffset );
        void    FlllColumnDisplayName();
        void    FillColumnProperties( const String& Dbms );
        BYTE*   GetDataPointerInCache( size_t idxRow, size_t idxColumn );
        SQLLEN  GetIndicatorValue( size_t idxRow, size_t idxColumn );
        void    UnbindAndDeleteCache();
        void    ZeroCache( size_t idxOffset );

        // Data members
        size_t          ROW_ARRAY_SIZE;
        size_t          MAX_COLUMN_LENGTH_BYTES;
        size_t          MAX_CACHE_RECORDS;
        size_t          m_uNumColumns;
        size_t          m_uCacheSize;
        size_t          m_uCacheRowSizeBytes;
        size_t          m_idxCacheStart;
        size_t          m_uNumRowsCached;
        size_t          m_idxCurrentRow;
        SQLUINTEGER     m_uCursorType;
        SQLUINTEGER     m_uCursorConcurrency;
        String          m_FieldValue;
        String          m_ColumnName;
        Odbc            m_Odbc;
        OdbcDatabase*   m_pDatabase;
        BYTE*           m_pCache;
        SQLHSTMT        m_StatementHandle;
        TYPE_COLUMN_PROPERTIES* m_pColumns;
};

#endif  // WINAUDIT_ODBC_RECORD_SET_H_
