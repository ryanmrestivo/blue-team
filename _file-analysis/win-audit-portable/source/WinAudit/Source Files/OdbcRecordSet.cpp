///////////////////////////////////////////////////////////////////////////////////////////////////
//
// ODBC Record Set Class Implementation
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

// A minimal record set object.

// Simple testing by executing a statement and stopping the database shows that
// most ODBC functions that describe results do not entail a round trip to the
// database. However, if the statement is only prepared then then trips to the
// database are usually required. Of course, observed behaviour is subject
// change. Y = trip to database, N = no trip to database required.
//                        Access            MySQL          SQL Server
//                  Prepared Executed Prepared Executed Prepared Executed
// SQLColAttribute      Y        N        Y        N        Y        N
// SQLDescribeCol       Y        N        Y        N        Y        N
// SQLGetDescField      Y        N        Y        N        Y        N
// SQLGetTypeInfo       N        Y        N        N        Y        Y
// SQLNumResultCols     N        N        Y        N        Y        N
// SQLNumParams         N        N        N        N        Y        N

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/OdbcRecordSet.h"

// 2. C System Files
#include <sqlext.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/BoundsException.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/SystemException.h"

// 5. This Project
#include "WinAudit/Header Files/OdbcDatabase.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
OdbcRecordSet::OdbcRecordSet()
              :ROW_ARRAY_SIZE( 50 ),
               MAX_COLUMN_LENGTH_BYTES( 0x1000 ),
               MAX_CACHE_RECORDS( 0xFFFF ),
               m_uNumColumns( 0 ),
               m_uCacheSize( ROW_ARRAY_SIZE ),
               m_uCacheRowSizeBytes( 0 ),
               m_idxCacheStart( 0 ),
               m_uNumRowsCached( 0 ),
               m_idxCurrentRow( 0 ),
               m_uCursorType( SQL_CURSOR_FORWARD_ONLY ),
               m_uCursorConcurrency( SQL_CONCUR_READ_ONLY ),
               m_FieldValue(),
               m_ColumnName(),
               m_Odbc(),
               m_pDatabase( nullptr ),
               m_pCache( nullptr ),
               m_StatementHandle( nullptr ),
               m_pColumns( nullptr )
{
}

// Copy constructor - not allowed so no implementation

// Destructor
OdbcRecordSet::~OdbcRecordSet()
{
    // Clean up
    try
    {
        Close();
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
//      Close the record set
//
//  Parameters:
//      None
//
//  Remarks:
//      Deletes the statement object and purges the cache
//
//  Returns:
//      void
//===============================================================================================//
void OdbcRecordSet::Close()
{
    // Delete the caches, do the values first because need to delete any
    // allocated strings. The values array is the same width as the columns
    // array so if were to delete the columns first then would no longer know
    // the array width, i.e. m_uNumColumns = 0
    UnbindAndDeleteCache();
    DeleteColumns();
    if ( m_StatementHandle )
    {
        m_Odbc.FreeStmt( m_StatementHandle, SQL_UNBIND );
        m_Odbc.FreeStmt( m_StatementHandle, SQL_RESET_PARAMS );
        m_Odbc.FreeStmt( m_StatementHandle, SQL_CLOSE );
        m_Odbc.FreeHandle( SQL_HANDLE_STMT, m_StatementHandle );
        m_StatementHandle = nullptr;     // Reset
    }
}

//===============================================================================================//
//  Description:
//      Get the value of the field at the specified column in the current
//      record
//
//  Parameters:
//      idxColumn - the zero-based index of the column in the query
//
//  Remarks:
//      Rows are cached, so can get the cached value, otherwise will
//      move to the record.
//
//  Returns:
//      Constant reference to a string
//===============================================================================================//
const String& OdbcRecordSet::FieldValue( size_t idxColumn  )
{
    return FieldValue( m_idxCurrentRow, idxColumn );
}

//===============================================================================================//
//  Description:
//      Get the value of the field at the specified row and column
//
//  Parameters:
//      idxRow    - the zero-based index of the row in the query
//      idxColumn - the zero-based index of the column in the query
//
//  Returns:
//      Constant reference to a string
//===============================================================================================//
const String& OdbcRecordSet::FieldValue( size_t idxRow, size_t idxColumn )
{
    BYTE* pBuffer;
    TYPE_COLUMN_PROPERTIES* pColumn = nullptr;

    // Get the column
    if ( m_pColumns == nullptr )
    {
        throw FunctionException( L"m_pColumns", __FUNCTION__ );
    }

    if ( idxColumn >= m_uNumColumns )
    {
        throw BoundsException( L"idxColumn", __FUNCTION__ );
    }
    pColumn = m_pColumns + idxColumn;

    // Test for a deleted row
    if ( IsDeleted( idxRow ) )
    {
        m_FieldValue = L"[Deleted]";
        return m_FieldValue;
    }

    pBuffer = GetDataPointerInCache( idxRow, idxColumn );
    DataBufferToString( pBuffer, pColumn );

    return m_FieldValue;
}

//===============================================================================================//
//  Description:
//      Get the number of columns in the record set
//
//  Parameters:
//      None
//
//  Returns:
//      unsigned integer of the number of columns
//===============================================================================================//
size_t OdbcRecordSet::GetColumnCount() const
{
    return m_uNumColumns;
}

//===============================================================================================//
//  Description:
//      Get the name of a column for visible display
//
//  Parameters:
//      idxColumn - the zero-based index of the column in the query
//
//  Returns:
//      Constant pointer to the column name
//===============================================================================================//
LPCWSTR OdbcRecordSet::GetColumnDisplayName( size_t idxColumn  ) const
{
    if ( ( m_pColumns == nullptr ) || ( m_uNumColumns == 0 ) )
    {
       throw FunctionException( L"m_pColumns/m_uNumColumns", __FUNCTION__ );
    }

    if ( idxColumn >= m_uNumColumns )
    {
        throw BoundsException( L"idxColumn", __FUNCTION__ );
    }

    return m_pColumns[ idxColumn ].szDisplayName;
}

//===============================================================================================//
//  Description:
//      Get the maximum number of characters required to display data for the
//      specified columns.
//
//  Parameters:
//      idxColumn - the zero-based index of the column in the query
//
//  Returns:
//      Number of characters required to display the data from a column
//===============================================================================================//
size_t OdbcRecordSet::GetColumnDisplaySizeChars( size_t idxColumn ) const
{
    if ( ( m_pColumns == nullptr ) || ( m_uNumColumns == 0 ) )
    {
       throw FunctionException( L"m_pColumns/m_uNumColumns", __FUNCTION__ );
    }

    if ( idxColumn >= m_uNumColumns )
    {
        throw BoundsException( L"idxColumn", __FUNCTION__ );
    }

    return m_pColumns[ idxColumn ].displayChars;
}

//===============================================================================================//
//  Description:
//      Get the name of a column in the record set
//
//  Parameters:
//      idxColumn      - the zero-based index of the column in the query
//      tableDotColumn - flag that indicates if want table.column form
//
//  Returns:
//      Constant reference to the column name
//===============================================================================================//
const String& OdbcRecordSet::GetColumnName( size_t idxColumn,
                                            bool tableDotColumn )
{
    Formatter Format;

    m_ColumnName = PXS_STRING_EMPTY;
    if ( ( m_pColumns == nullptr ) || ( m_uNumColumns == 0 ) )
    {
       throw FunctionException( L"m_pColumns/m_uNumColumns", __FUNCTION__ );
    }

    if ( idxColumn >= m_uNumColumns )
    {
        throw BoundsException( L"idxColumn", __FUNCTION__ );
    }

    // See if want table.name form use that in preference
    if ( tableDotColumn )
    {
        m_ColumnName = m_pColumns[ idxColumn ].szTable;
        m_ColumnName.Trim();
        if ( m_ColumnName.GetLength() )
        {
            m_ColumnName += PXS_CHAR_DOT;
        }
        m_ColumnName += m_pColumns[ idxColumn ].szName;
    }

    // Next, try the name
    if ( m_ColumnName.IsEmpty() )
    {
        m_ColumnName += m_pColumns[ idxColumn ].szName;
        m_ColumnName.Trim();
    }

    // If sill no name, use the column number
    if ( m_ColumnName.IsEmpty() )
    {
        m_ColumnName  = L"Column_";
        m_ColumnName += Format.SizeT( PXSAddSizeT( idxColumn, 1 ) );
    }

    return m_ColumnName;
}

//===============================================================================================//
//  Description:
//      Get the precision of a column
//
//  Parameters:
//      idxColumn - the zero-based index of the column in the query
//
//  Returns:
//      unsigned char of the column's precision
//===============================================================================================//
SQLCHAR OdbcRecordSet::GetColumnPrecision( size_t idxColumn ) const
{
    if ( ( m_pColumns == nullptr ) || ( m_uNumColumns == 0 ) )
    {
       throw FunctionException( L"m_pColumns/m_uNumColumns", __FUNCTION__ );
    }

    if ( idxColumn >= m_uNumColumns )
    {
        throw BoundsException( L"idxColumn", __FUNCTION__ );
    }

    return m_pColumns[ idxColumn ].bPrecision;
}

//===============================================================================================//
//  Description:
//      Get the scale for a column, used for SQL_DECIMAL and SQL_NUMERIC types
//
//  Parameters:
//      idxColumn - the zero-based index of the column in the query
//
//  Returns:
//      SQLSCHAR of the scale, zero for non NUMERIC/DECIMAL columns
//===============================================================================================//
SQLSCHAR OdbcRecordSet::GetColumnScale( size_t idxColumn ) const
{
    SQLSCHAR scale = 0;

    if ( ( m_pColumns == nullptr ) || ( m_uNumColumns == 0 ) )
    {
       throw FunctionException( L"m_pColumns/m_uNumColumns", __FUNCTION__ );
    }

    if ( idxColumn >= m_uNumColumns )
    {
        throw BoundsException( L"idxColumn", __FUNCTION__ );
    }

    if ( ( m_pColumns[ idxColumn ].sqlType == SQL_NUMERIC ) ||    // 2
         ( m_pColumns[ idxColumn ].sqlType == SQL_DECIMAL )  )    // 3
    {
        scale = m_pColumns[ idxColumn ].scale;
    }

    return scale;
}

//===============================================================================================//
//  Description:
//      Get the size of a column in bytes
//
//  Parameters:
//      idxColumn - the zero-based index of the column in the query
//
//  Returns:
//      The column size in bytes
//===============================================================================================//
size_t OdbcRecordSet::GetColumnSizeBytes( size_t idxColumn ) const
{
    if ( ( m_pColumns == nullptr ) || ( m_uNumColumns == 0 ) )
    {
       throw FunctionException( L"m_pColumns/m_uNumColumns", __FUNCTION__ );
    }

    if ( idxColumn >= m_uNumColumns )
    {
        throw BoundsException( L"idxColumn", __FUNCTION__ );
    }

    return m_pColumns[ idxColumn ].sizeBytes;
}

//===============================================================================================//
//  Description:
//      Get the SQL data type of a column
//
//  Parameters:
//      idxColumn - the zero-based index of the column in the query
//
//  Returns:
//      signed integer of the data type, e.g. SQL_INTEGER
//===============================================================================================//
SQLSMALLINT OdbcRecordSet::GetColumnSqlType( size_t idxColumn ) const
{
    if ( ( m_pColumns == nullptr ) || ( m_uNumColumns == 0 ) )
    {
       throw FunctionException( L"m_pColumns/m_uNumColumns", __FUNCTION__ );
    }

    if ( idxColumn >= m_uNumColumns )
    {
        throw BoundsException( L"idxColumn", __FUNCTION__ );
    }

    return m_pColumns[ idxColumn ].sqlType;
}

//===============================================================================================//
//
//  Description:
//      Get if a the specified row has been deleted
//
//  Parameters:
//      idxRow - the zero-based index of the row in the query
//
//  Remarks:
//      If the record has been deleted then the indicator value of the
//      first column has been set to SQL_ROW_DELETED
//
//  Returns:
//      true if the record has been deleted, else false
//===============================================================================================//
bool OdbcRecordSet::IsDeleted( size_t idxRow )
{
    if ( SQL_ROW_DELETED == GetIndicatorValue( idxRow, 0 ) )  // zeroth column
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Get if a field is NULL
//
//  Parameters:
//      idxColumn - the zero-based index of the column in the query
//      idxRow - the zero-based index of the row in the query
//
//  Remarks:
//      If the field is NULL, its indicator value is SQL_NULL_DATA
//
//  Returns:
//      true if NULL otherwise false
//===============================================================================================//
bool OdbcRecordSet::IsFieldNull( size_t idxRow, size_t idxColumn )
{
    if ( SQL_NULL_DATA == GetIndicatorValue( idxRow, idxColumn ) )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Get if the field at the current row and specified column is NULL
//
//  Parameters:
//      idxColumn - the zero-based index of the column in the query
//
//  Returns:
//      true if NULL otherwise false
//===============================================================================================//
bool OdbcRecordSet::IsFieldNull( size_t idxColumn )
{
    return IsFieldNull( m_idxCurrentRow, idxColumn );
}

//===============================================================================================//
//  Description:
//      Determine if the record set is open
//
//  Parameters:
//      None
//
//  Returns:
//      true if the record set is open otherwise false
//===============================================================================================//
bool OdbcRecordSet::IsOpen() const
{
    // If have created a columns array the record set is open
    if ( m_pColumns && m_uNumColumns )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Move to a specified record in the record set
//
//  Parameters:
//      idxRow - the number of the row, counting starts at 0.
//
//  Returns:
//      true on success, otherwise false
//===============================================================================================//
bool OdbcRecordSet::Move( size_t idxRow )
{
    bool   success     = false;
    size_t idxTemp     = 0, idxNewStart = 0;
    SQLLEN fetchOffset = 0;

    // Class scope checks
    if ( m_uCacheSize == 0 )
    {
        return false;
    }

    // Handle the special case of no rows yet fetched
    if ( ( idxRow == 0 )           &&
         ( m_idxCurrentRow  == 0 ) &&
         ( m_uNumRowsCached == 0 )  )
    {
        // Do the first fetch, can always use SQL_FETCH_NEXT regardless
        // of cursor type
        if ( FillCache( SQL_FETCH_NEXT , 0 ) )
        {
            success = true;
            m_idxCacheStart = 0;
            m_idxCurrentRow = 0;
        }
        return success;
    }

    // Test if the specified row is in the current rowset
    if ( ( idxRow >= m_idxCacheStart ) &&
         ( idxRow  < PXSAddSizeT( m_idxCacheStart, m_uCacheSize ) ) )
    {
        // The cache may not be completely filled if on the last rowset so
        // test if in bounds of the rows that are cached. Do not want to
        // fetch any more records.
        if ( idxRow < PXSAddSizeT( m_idxCacheStart, m_uNumRowsCached ) )
        {
            success = true;
            m_idxCurrentRow = idxRow;
        }

        return success;
    }

    // Fetch a new rowset, test if using a forward only cursor
    // otherwise using a scrollable one
    if ( SQL_CURSOR_FORWARD_ONLY == m_uCursorType )
    {
        // The desired row must be in the next rowset.
        idxNewStart = PXSAddSizeT( m_idxCacheStart, m_uCacheSize );
        if ( ( idxRow < idxNewStart ) ||
             ( idxRow > PXSAddSizeT( idxNewStart, m_uCacheSize ) ) )
        {
            throw BoundsException( L"idxRow", __FUNCTION__ );
        }

        if ( FillCache( SQL_FETCH_NEXT, 0 ) )
        {
            // Set the new cache start and the current row
            m_idxCacheStart = idxNewStart;
            m_idxCurrentRow = m_idxCacheStart;

            // Verify that the specified row is in bounds of those
            // cached as could have fetched the last rowset. For
            // example, if rowset size is 10 and only fetched rows
            // 10 - 15 but wanted 19, so have failed to move.
            if ( idxRow < PXSAddSizeT( m_idxCacheStart, m_uNumRowsCached ) )
            {
                success = true;
                m_idxCurrentRow = idxRow;
            }
        }
    }
    else
    {
        // Can use a scrollable cursor
        idxTemp     = PXSAddSizeT( idxRow, 1 );
        fetchOffset = PXSCastSizeTToSqlLen( idxTemp );
        if ( FillCache( SQL_FETCH_ABSOLUTE, fetchOffset ) )
        {
            success = true;
            m_idxCacheStart = idxRow;
            m_idxCurrentRow = idxRow;
        }
    }

    return success;
}

//===============================================================================================//
//  Description:
//      Move to the next record in the record set
//
//  Parameters:
//      None
//
//  Returns:
//      true on success, otherwise false
//===============================================================================================//
bool OdbcRecordSet::MoveNext()
{
    size_t idxRow = 0;

    // Handle special case of record set opened, but no records yet fetched
    if ( ( m_idxCurrentRow  == 0 ) && ( m_uNumRowsCached == 0 ) )
    {
        idxRow = 0;
    }
    else
    {
        idxRow = PXSAddSizeT( m_idxCurrentRow, 1 );
    }
    return Move( idxRow );
}

//===============================================================================================//
//  Description:
//      Open a record set
//
//  Parameters:
//      SqlQuery  - the SQL statement
//      pDatabase - pointer to database object
//
//  Remarks:
//      Do not use a table alias in the query if need to know the actual table
//      name. MySQL returns this alias when use SQLColAttribute with
//      SQL_DESC_TABLE_NAME and SQL_DESC_BASE_TABLE_NAME returns an empty
//      string so cannot get the table name to which a column belongs.
//
//  Returns:
//      void
//===============================================================================================//
void OdbcRecordSet::Open( const String& SqlQuery, OdbcDatabase* pDatabase )
{
    DWORD    queryTimeOut = 0;
    String   Dbms;
    SQLHDBC  hDBC = nullptr;

    // See if its returning results, allow '{' because it is used by
    // ODBC to call procedures
    if ( ( SqlQuery.StartsWith( L"SELECT", false ) == false ) &&
         ( SqlQuery.StartsWith( L"SHOW"  , false ) == false ) &&
         ( SqlQuery.StartsWith( L"{"     , false ) == false )  )
    {
        throw ParameterException( L"SqlQuery", __FUNCTION__ );
    }

    if ( pDatabase == nullptr )
    {
        throw ParameterException( L"pDB", __FUNCTION__ );
    }

    // Need the connection handle
    hDBC  = pDatabase->GetConnectionHandle();
    if ( hDBC == nullptr )
    {
        throw NullException( L"hDBC", __FUNCTION__ );
    }

    // Free resources, order is important so will delete any strings
    Close();
    UnbindAndDeleteCache();
    DeleteColumns();
    m_uNumColumns    = 0;
    m_idxCacheStart  = 0;
    m_uNumRowsCached = 0;
    m_idxCurrentRow  = 0;
    m_pColumns       = nullptr;
    m_pCache         = nullptr;

    // Allocate the statement handle
    m_pDatabase = pDatabase;
    m_Odbc.AllocHandle( SQL_HANDLE_STMT, hDBC, &m_StatementHandle );

    // Set the query time out and other properties
    if ( m_pDatabase->SupportsQueryTimeOut() )
    {
        queryTimeOut = pDatabase->QueryTimeOut();
        m_Odbc.SetStmtAttr( m_StatementHandle,
                            SQL_ATTR_QUERY_TIMEOUT,
                            (SQLPOINTER)(ULONG_PTR)queryTimeOut,    // TYPE CAST
                            0 );
    }
    m_Odbc.SetStmtAttr( m_StatementHandle,
                        SQL_ATTR_CURSOR_TYPE,
                        (SQLPOINTER)(ULONG_PTR)m_uCursorType,       // TYPE CAST
                        0 );
    m_Odbc.SetStmtAttr(
                    m_StatementHandle,
                    SQL_ATTR_CONCURRENCY,
                    (SQLPOINTER)(ULONG_PTR)m_uCursorConcurrency,    // TYPE CAST
                    0 );
    m_Odbc.SetStmtAttr( m_StatementHandle,
                        SQL_ATTR_ROW_ARRAY_SIZE,
                        (SQLPOINTER)(ULONG_PTR)ROW_ARRAY_SIZE,      // TYPE CAST
                        0 );
    Dbms = m_pDatabase->GetDbmsName();

    // Execute and fill the meta-data
    m_Odbc.ExecDirect( m_StatementHandle, SqlQuery );
    FillColumnProperties( Dbms );
    FlllColumnDisplayName();
    AllocateCache();
    BindCache();
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Allocate memory for the results and bind it to the statement
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void OdbcRecordSet::AllocateCache()
{
    size_t i = 0, bytesAllocated = 0, numBytes = 0, rounding = 0;

    // First ensure the cache is deleted
    UnbindAndDeleteCache();

    if ( ( m_pColumns == nullptr ) || ( m_uNumColumns == 0 ) )
    {
       throw FunctionException( L"m_pColumns/m_uNumColumns", __FUNCTION__ );
    }

    if ( m_pDatabase == nullptr )
    {
        throw FunctionException( L"m_pDatabase", __FUNCTION__ );
    }

    if ( m_uCacheSize == 0 )
    {
        throw FunctionException( L"m_uCacheSize", __FUNCTION__ );
    }

    if ( m_StatementHandle == nullptr )
    {
        throw FunctionException( L"m_StatementHandle", __FUNCTION__ );
    }

    // Determine how many bytes will be allocated for each column
    for ( i = 0; i < m_uNumColumns; i++ )
    {
       bytesAllocated = 0;
        switch ( m_pColumns[ i ].sqlType )
        {
            // Fall through
            default:
            case SQL_UNKNOWN_TYPE:      // 0

                throw SystemException( ERROR_INVALID_DATA, L"SQL_UNKNOWN_TYPE", __FUNCTION__ );

            // Fall through - fixed length deta
            case SQL_NUMERIC:           // 2
            case SQL_DECIMAL:           // 3
            case SQL_INTEGER:           // 4
            case SQL_SMALLINT:          // 5
            case SQL_FLOAT:             // 6 - this is a double
            case SQL_REAL:              // 7 - this is single precision
            case SQL_DOUBLE:            // 8
            case SQL_TYPE_DATE:         // 91   ODBC 3.0
            case SQL_TYPE_TIME:         // 92   ODBC 3.0
            case SQL_TYPE_TIMESTAMP:    // 93   ODBC 3.0
            case SQL_BIGINT:            // (-5)
            case SQL_TINYINT:           // (-6)
            case SQL_BIT:               // (-7)
            case SQL_GUID:              // (-11)
                bytesAllocated = m_pColumns[ i ].sizeBytes;
                break;

            // Variable length data - ANSI strings.
            case SQL_CHAR:              // 1
            case SQL_VARCHAR:           // 12
            case SQL_LONGVARCHAR:       // (-1)

                bytesAllocated = PXSAddSizeT( m_pColumns[ i ].sizeBytes,
                                              sizeof ( SQLCHAR ) );  // + NULL
                PXSLimitSizeT( sizeof ( SQLCHAR ), MAX_COLUMN_LENGTH_BYTES, &bytesAllocated );
                break;

            // Variable length data - wide strings.
            case SQL_WCHAR:             // (-8)
            case SQL_WVARCHAR:          // (-9)
            case SQL_WLONGVARCHAR:      // (-10)
                bytesAllocated = PXSAddSizeT( m_pColumns[ i ].sizeBytes,
                                              sizeof ( SQLWCHAR ) );  // + NULL
                if ( ( bytesAllocated % 2 ) == 1 )
                {
                    bytesAllocated = PXSAddSizeT( bytesAllocated, 1 );
                }
                PXSLimitSizeT( sizeof ( SQLWCHAR ), MAX_COLUMN_LENGTH_BYTES, &bytesAllocated );
                break;

            // Variable length data - binary data
            case SQL_BINARY:            // (-2)
            case SQL_VARBINARY:         // (-3)
            case SQL_LONGVARBINARY:     // (-4)
                bytesAllocated = m_pColumns[ i ].sizeBytes;
                PXSLimitSizeT( 1, MAX_COLUMN_LENGTH_BYTES, &bytesAllocated );
                break;
        }

        // Do not allocate too many bytes
        if ( ( bytesAllocated == 0 ) ||
             ( bytesAllocated > MAX_COLUMN_LENGTH_BYTES ) )
        {
            throw SystemException( ERROR_INVALID_DATA, L"bytesAllocated", __FUNCTION__ );
        }
        m_pColumns[ i ].bytesAllocated = bytesAllocated;
    }

    // Compute the width of a row
    m_uCacheRowSizeBytes = 0;
    for ( i = 0; i < m_uNumColumns; i++ )
    {
        m_uCacheRowSizeBytes = PXSAddSizeT( m_uCacheRowSizeBytes, m_pColumns[ i ].bytesAllocated );

        // The indicator column
        m_uCacheRowSizeBytes = PXSAddSizeT( m_uCacheRowSizeBytes, sizeof ( SQLLEN ) );
    }

    // Round up to pointer size
    size_t ptrSize = sizeof ( void* );
    if ( m_uCacheRowSizeBytes % ptrSize )
    {
        rounding = ptrSize - ( m_uCacheRowSizeBytes % ptrSize );
    }
    m_uCacheRowSizeBytes = PXSAddSizeT( m_uCacheRowSizeBytes, rounding );

    numBytes = PXSMultiplySizeT( m_uCacheRowSizeBytes, m_uCacheSize );
    m_pCache = new BYTE[ numBytes ];
    if ( m_pCache == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    memset( m_pCache, 0, numBytes );
}

//===============================================================================================//
//  Description:
//      Bind the cache that was allocated in AllocateCache
//
//  Parameters:
//      None
//
//  Remarks:
//      Using row-wise binding, documentation says that this is somewhat
//      faster than column-wise binding and allows use of
//      SQL_ATTR_ROW_BIND_OFFSET_PTR so that can fetch records incrementally
//      to fill the cache.
//
//  Returns:
//      void
//===============================================================================================//
void OdbcRecordSet::BindCache()
{
    String    Dbms;
    size_t    i = 0, offset = 0;
    SQLLEN    BufferLength  = 0;
    SQLLEN*   pStrLen_or_Ind   = nullptr;
    SQLHDESC  DescriptorHandle = nullptr;
    SQLSMALLINT  RecNumber     = 0;
    SQLUSMALLINT ColumnNumber  = 0;

    // Verify have the cache
    if ( ( m_pCache    == nullptr ) ||
         ( m_pColumns  == nullptr ) ||
         ( m_pDatabase == nullptr )  )
    {
        throw FunctionException( L"m_pCache/m_pColumns", __FUNCTION__ );
    }
    Dbms = m_pDatabase->GetDbmsName();

    // Compute the offsets of the columns and their indicators
    for ( i = 0; i < m_uNumColumns; i++ )
    {
        // Increment by column width
        m_pColumns[ i ].rowOffsetBytes = offset;
        offset = PXSAddSizeT( offset, m_pColumns[ i ].bytesAllocated );

        // Increment by the indicator's size
        m_pColumns[ i ].indRowOffsetBytes = offset;
        offset = PXSAddSizeT( offset,  sizeof ( SQLLEN ) );
    }

    // Bind, using row-wise binding
    Dbms = m_pDatabase->GetDbmsName();
    m_Odbc.FreeStmt( m_StatementHandle, SQL_UNBIND );
    m_Odbc.SetStmtAttr( m_StatementHandle,
                        SQL_ATTR_ROW_BIND_TYPE,
                        (SQLPOINTER)(ULONG_PTR)m_uCacheRowSizeBytes,  // TYPE CAST
                    0 );
    for ( i = 0; i < m_uNumColumns; i++ )
    {
        ColumnNumber = PXSCastSizeTToUInt16( i + 1 );

        // Type cast the buffer length to SQLLEN
        if ( m_pColumns[ i ].bytesAllocated > PXS_SQLLEN_MAX )
        {
            throw BoundsException( L"bytesAllocated", __FUNCTION__ );
        }
        BufferLength = PXSCastSizeTToSqlLen( m_pColumns[i].bytesAllocated );

        // pIndicator is where the SQLLEN indicator is located
        pStrLen_or_Ind = reinterpret_cast<SQLLEN*>(
                                   m_pCache + m_pColumns[i].indRowOffsetBytes );
        m_Odbc.BindCol( m_StatementHandle,
                        ColumnNumber,
                        m_pColumns[ i ].cType,
                        m_pCache + m_pColumns[ i ].rowOffsetBytes,
                        BufferLength, pStrLen_or_Ind );

        // For numeric columns the default precision is ODBC driver dependent
        // and the scale is set to zero. This means decimal information can
        // be lost when the data is fetched into a SQL_NUMERIC_STRUCT. This
        // happens with Access and SQL Server. So, can either bind numeric
        // data to strings or follow article Q181254. The latter is done here.
        // Note, must do this after binding the column. Ignore Oracle, if set
        // any of these descriptors, then get no data when do a fetch.
        if ( ( SQL_C_NUMERIC == m_pColumns[ i ].cType  ) &&
             ( Dbms.CompareI( PXS_DBMS_NAME_ORACLE ) == 0 ) )
        {
            RecNumber = PXSCastSizeTToInt16( i + i );
            DescriptorHandle      = nullptr;
            m_Odbc.GetStmtAttr( m_StatementHandle,
                                SQL_ATTR_APP_ROW_DESC, &DescriptorHandle, 0, nullptr );
            if ( DescriptorHandle )
            {
                // Set the data type
                m_Odbc.SetDescField( DescriptorHandle,
                                     RecNumber,
                                     SQL_DESC_TYPE,
                                     reinterpret_cast<SQLPOINTER>( SQL_C_NUMERIC ),
                                     0 );          // Not used

                // Precision
                m_Odbc.SetDescField(
                                  DescriptorHandle,
                                  RecNumber,
                                  SQL_DESC_PRECISION,
                                  (SQLPOINTER)(ULONG_PTR)m_pColumns[i].bPrecision,  // TYPE CAST
                                  0 );

                // Scale
                m_Odbc.SetDescField( DescriptorHandle,
                                     RecNumber,
                                     SQL_DESC_SCALE,
                                     (SQLPOINTER)(ULONG_PTR)m_pColumns[ i ].scale,  // TYPE CAST
                                     0 );

                // Set the data pointer
                m_Odbc.SetDescField(
                        DescriptorHandle,
                        RecNumber,
                        SQL_DESC_DATA_PTR,
                        reinterpret_cast<SQLPOINTER>(m_pCache +  m_pColumns[i].rowOffsetBytes),
                        0 );
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Convert the data to a string
//
//  Parameters:
//      pBuffer - pointer into the row cache
//      pColumn - pointer to the column properties
//
//  Remarks:
//      The string in stored at class scope
//
//  Returns:
//      void
//===============================================================================================//
void OdbcRecordSet::DataBufferToString( BYTE* pBuffer, TYPE_COLUMN_PROPERTIES* pColumn )
{
    BYTE   scale = 0;
    WORD   year  = 0;
    double temp  = 0.0;
    char*     pszAnsi = nullptr;
    wchar_t*  pwzWide = nullptr;
    Formatter Format;
    SQL_DATE_STRUCT* pDate = nullptr;
    SQL_TIME_STRUCT* pTime = nullptr;
    SQL_NUMERIC_STRUCT*   pNumeric   = nullptr;
    SQL_TIMESTAMP_STRUCT* pTimeStamp = nullptr;

    if ( pColumn == nullptr )
    {
        throw NullException( L"pColumn", __FUNCTION__ );
    }

    // Test for NULL
    if ( pBuffer == nullptr )
    {
        m_FieldValue = L"Null";
        return;
    }

    // Test for special formatting
    if ( pColumn->boolean )
    {
        if ( *reinterpret_cast<BYTE*>( pBuffer ) == 0 )
        {
            m_FieldValue = L"0";
        }
        else
        {
            m_FieldValue = L"1";
        }
        return;
    }

    // Format according to data type
    switch( pColumn->sqlType )
    {
        // Fall through
        default:
        case SQL_UNKNOWN_TYPE:    // 0
            m_FieldValue = L"[UNKNOWN]";
            break;

        case SQL_CHAR:            // 1
        case SQL_VARCHAR:         // 12
        case SQL_LONGVARCHAR:     // (-1)

            if ( pColumn->bytesAllocated )
            {
                pszAnsi = reinterpret_cast<char*>( pBuffer );
                pszAnsi[ pColumn->bytesAllocated - 1 ] = PXS_CHAR_NULL;
                m_FieldValue.SetAnsi( pszAnsi );
            }
            break;

        case SQL_WCHAR:           // (-8)
        case SQL_WVARCHAR:        // (-9)
        case SQL_WLONGVARCHAR:    // (-10)

            if ( pColumn->bytesAllocated >= sizeof ( SQLWCHAR ) )
            {
              pwzWide = reinterpret_cast<wchar_t*>( pBuffer );
              pwzWide[( pColumn->bytesAllocated / sizeof (SQLWCHAR) ) - 1 ] = 0;
              m_FieldValue = pwzWide;
            }
            break;

        // Fall through
        case SQL_NUMERIC:         // 2
        case SQL_DECIMAL:         // 3

            pNumeric = reinterpret_cast<SQL_NUMERIC_STRUCT*>( pBuffer );
            temp     = m_Odbc.SqlNumericStructToDouble( pNumeric );
            scale    = PXSCastInt8ToUInt8( pNumeric->scale );
            m_FieldValue = Format.Double( temp, scale );
            break;

        case SQL_SMALLINT:        // 4

            m_FieldValue = Format.Int32( *reinterpret_cast<SQLSMALLINT*>( pBuffer ) );
            break;

        case SQL_INTEGER:         // 5

            m_FieldValue = Format.Int32( *reinterpret_cast<SQLINTEGER*>( pBuffer ) );
            break;

        case SQL_FLOAT:           // 6

            // SQL_FLOAT can be variable precision, if its 24 = SQL_REAL and
            // 53 = SQL_DOUBLE. However, in the SQL header is defined
            // as a double
            m_FieldValue = Format.Double( *reinterpret_cast<SQLDOUBLE*>( pBuffer ) );
            break;

        case SQL_REAL:            // 7

            m_FieldValue = Format.Float( *reinterpret_cast<SQLREAL*> (pBuffer));
            break;

        case SQL_DOUBLE:          // 8

            m_FieldValue = Format.Double( *reinterpret_cast<SQLDOUBLE*>( pBuffer ) );
            break;

        case SQL_TYPE_DATE:       // 91   ODBC 3.0

            // The year member of DATE_STRUCT is a SQLSMALLINT but want a WORD
            pDate = reinterpret_cast<SQL_DATE_STRUCT*>( pBuffer );
            year  = PXSCastInt16ToUInt16( pDate->year );
            m_FieldValue = Format.DateInUserLocale( year, pDate->month, pDate->day );
            break;

        case SQL_TYPE_TIME:         // 92   ODBC 3.0

            pTime = reinterpret_cast<SQL_TIME_STRUCT*>( pBuffer );
            m_FieldValue = Format.TimeInUserLocale( pTime->hour, pTime->minute, pTime->second );
            break;

        case SQL_TYPE_TIMESTAMP:    // 93   ODBC 3.0

            // The year member of DATE_STRUCT is a SQLSMALLINT but want a WORD
            // Access stores data and time as timestamps so might not get the
            // expected formatting
            pTimeStamp = reinterpret_cast<SQL_TIMESTAMP_STRUCT*>( pBuffer );
            year = PXSCastInt16ToUInt16( pTimeStamp->year );
            m_FieldValue = Format.DateInUserLocale( year, pTimeStamp->month, pTimeStamp->day );
            m_FieldValue += PXS_CHAR_SPACE;
            m_FieldValue += Format.TimeInUserLocale( pTimeStamp->hour,
                                                    pTimeStamp->minute, pTimeStamp->second );
            break;

        // Fall through
        case SQL_BINARY:          // (-2)
        case SQL_VARBINARY:       // (-3)
        case SQL_LONGVARBINARY:   // (-4)

            m_FieldValue = L"[Binary]";
            break;

        case SQL_BIGINT:          // (-5)

            m_FieldValue = Format.Int64( *reinterpret_cast<SQLBIGINT*>( pBuffer ) );
            break;

        case SQL_TINYINT:         // (-6)

            m_FieldValue = Format.Int32( *reinterpret_cast<SQLSCHAR*>( pBuffer ) );
            break;

        case SQL_BIT:             // (-7)

            m_FieldValue = Format.UInt8Hex( *reinterpret_cast<SQLCHAR*>(pBuffer), false );
            break;

        case SQL_GUID:            // (-11)   ODBC 3.5

            m_FieldValue = Format.GuidToString( *reinterpret_cast<SQLGUID*>( pBuffer ) );
            break;
    }
}

//===============================================================================================//
//  Description:
//      Clean up the array of columns
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void OdbcRecordSet::DeleteColumns()
{
    if ( m_pColumns )
    {
        delete[] m_pColumns;
        m_pColumns = nullptr;
    }
    m_uNumColumns = 0;
}

//===============================================================================================//
//  Description:
//      Fill the record set's cache
//
//  Parameters:
//      FetchOrientation - ODBC defined constant, e.g. SQL_FETCH_NEXT
//      FetchOffset      - offset for the fetch
//
//  Remarks:
//      see SQLFetchScroll()
//
//  Returns:
//      SQLULEN of the number of records fetched
//===============================================================================================//
size_t OdbcRecordSet::FillCache( SQLSMALLINT FetchOrientation, SQLLEN FetchOffset )
{
    size_t    rowBindOffset = 0, rowsFetched = 0, totalRowsFetched = 0;
    SQLULEN   attrRowBindOffset = 0, attrRowsFetched = 0;
    SQLRETURN sqlReturn = 0;

    // Verify have a statement
    if ( m_StatementHandle == nullptr )
    {
        throw FunctionException( L"m_StatementHandle", __FUNCTION__ );
    }

    // Set the statement's attributes
    m_Odbc.SetStmtAttr( m_StatementHandle, SQL_ATTR_ROWS_FETCHED_PTR, &attrRowsFetched, 0 );
    m_Odbc.SetStmtAttr( m_StatementHandle,
                        SQL_ATTR_ROW_ARRAY_SIZE,
                        (SQLPOINTER)(ULONG_PTR)ROW_ARRAY_SIZE,  // TYPE CAST
                        0 );      //
    m_Odbc.SetStmtAttr( m_StatementHandle,
                        SQL_ATTR_ROW_BIND_OFFSET_PTR, nullptr, 0 );
    sqlReturn = m_Odbc.FetchScroll( m_StatementHandle, FetchOrientation, FetchOffset );
    if ( SQL_SUCCEEDED( sqlReturn ) )
    {
        rowsFetched      = PXSCastSqlULenToSizeT( attrRowsFetched );
        totalRowsFetched = PXSAddSizeT( totalRowsFetched, rowsFetched );
    }

    // Fetch forward until got the required rows or end of rowset
    while ( ( SQL_SUCCEEDED( sqlReturn ) ) &&
            ( PXSAddSizeT( totalRowsFetched, ROW_ARRAY_SIZE ) <= m_uCacheSize) )
    {
        attrRowsFetched = 0;
        rowBindOffset = PXSMultiplySizeT( totalRowsFetched, m_uCacheRowSizeBytes );
        attrRowBindOffset = PXSCastSizeTToSqlULen( rowBindOffset );
        m_Odbc.SetStmtAttr( m_StatementHandle,
                            SQL_ATTR_ROW_BIND_OFFSET_PTR, &attrRowBindOffset, 0);
        sqlReturn = m_Odbc.Fetch( m_StatementHandle );

        // Tally
        if ( SQL_SUCCEEDED( sqlReturn ) )
        {
            totalRowsFetched = PXSAddSizeT( totalRowsFetched, attrRowsFetched);
        }
    }

    // Test for success
    if ( totalRowsFetched )
    {
        m_uNumRowsCached = totalRowsFetched;

        // Zero any unfilled rows in the cache
        if ( totalRowsFetched < m_uCacheSize )
        {
            ZeroCache( totalRowsFetched );
        }
    }

    return totalRowsFetched;
}

//===========================================================================//
//  Description:
//      Fill the display name for each column
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void OdbcRecordSet::FlllColumnDisplayName()
{
    size_t    i = 0 , j = 0 , count = 0;
    String    DisplayName;
    Formatter Format;

    if ( m_pColumns == nullptr )
    {
        throw FunctionException( L"m_pColumns", __FUNCTION__ );
    }

    // The precedence is "Name" then "table.column" if more than 1 table
    // or "column" if only one table.
    for ( i = 0; i < m_uNumColumns; i++ )
    {
        DisplayName = PXS_STRING_EMPTY;

        // Test if there are duplicate column names
        count = 0;
        if ( wcslen( m_pColumns[ i ].szName ) )
        {
            for ( j = 0; j < m_uNumColumns; j++ )
            {
                if ( CSTR_EQUAL == CompareString( LOCALE_INVARIANT,
                                                  NORM_IGNORECASE,
                                                  m_pColumns[ i ].szName,
                                                 -1,
                                                  m_pColumns[ j ].szName,
                                                  -1 ) )
                {
                    count++;
                }
            }
        }

        // If there are duplicate names use the table.column form
        if ( count > 1 )
        {
            DisplayName  = m_pColumns[ i ].szTable;
            DisplayName += PXS_CHAR_DOT;
            DisplayName += m_pColumns[ i ].szName;
        }
        else
        {
           DisplayName = m_pColumns[ i ].szName;
        }

        // If still no name, the use the column number
        if ( DisplayName.IsEmpty() )
        {
            DisplayName = Format.StringSizeT( L"Column_%%1", i );
        }

        if ( DisplayName.c_str() )
        {
            StringCchCopy( m_pColumns[ i ].szDisplayName,
                           ARRAYSIZE( m_pColumns[ i ].szDisplayName ), DisplayName.c_str() );
        }
    }
}

//===============================================================================================//
//  Description:
//      Fill the column properties array
//
//  Parameters:
//      Dbms - the database management system name
//
//  Remarks:
//      For performance reasons this should be done after a executing the
//      query. For example, with SQL Server SQLDescribeCol, SQLColAttribute
//      and SQLNumResultCols can cause round trips to the server.
//
//  Returns:
//      void
//===============================================================================================//
void OdbcRecordSet::FillColumnProperties( const String& Dbms )
{
    size_t       i = 0, numColumns = 0, sizeBytes = 0;
    SQLLEN       NumericAttribute  = 0;
    SQLULEN      ColumnSize = 0;
    SQLWCHAR     szBuffer[ 256 ] = { 0 };  // Identifier, size must be even
    SQLSMALLINT  ColumnCount   = 0, NameLength = 0, DataType = 0;
    SQLSMALLINT  DecimalDigits = 0, Nullable   = 0, StringLength = 0;
    SQLUSMALLINT ColumnNumber  = 0;

    // Must have an active statement
    if ( m_StatementHandle == nullptr )
    {
        throw NullException( L"m_StatementHandle", __FUNCTION__ );
    }

    m_Odbc.NumResultCols( m_StatementHandle, &ColumnCount );
    if ( ColumnCount <= 0 )     // Do not allow a sero column result set
    {
        throw BoundsException( L"sNumColumns", __FUNCTION__ );
    }
    numColumns = PXSCastInt16ToSizeT( ColumnCount );

    // Allocate columns
    DeleteColumns();
    m_pColumns = new TYPE_COLUMN_PROPERTIES[ numColumns ];
    if ( m_pColumns == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    memset( m_pColumns, 0, numColumns * sizeof (TYPE_COLUMN_PROPERTIES ) );
    m_uNumColumns = numColumns;

    // Add the column names from the query
    for ( i = 0; i < m_uNumColumns; i++ )
    {
        // Need a SQLSMALLINT
        if ( i >= UINT16_MAX )      // Inclusive as want to increment
        {
            throw BoundsException( L"i", __FUNCTION__ );
        }
        ColumnNumber = PXSCastSizeTToUInt16( i + 1 );

        NameLength    = 0;
        DataType      = 0;
        ColumnSize    = 0;
        DecimalDigits = 0;
        Nullable      = 0;
        m_Odbc.DescribeCol( m_StatementHandle,
                            ColumnNumber,
                            m_pColumns[ i ].szName,     // SQL_DESC_NAME
                            ARRAYSIZE( m_pColumns[ i ].szName ),
                            &NameLength,
                            &DataType,                  // SQL_DESC_CONCISE_TYPE
                            &ColumnSize,                // multiple meanings
                            &DecimalDigits,
                            &Nullable );
        m_pColumns[ i ].szName[ ARRAYSIZE( m_pColumns[ i ].szName ) - 1 ] = 0;
        m_pColumns[ i ].sqlType = DataType;

        // DecimalDigitsPtr is a SQLSMALLINT but SQL_NUMERIC_STRUCT
        // uses SQLSCHAR
        m_pColumns[ i ].scale  = PXSCastInt16ToInt8( DecimalDigits );

        // Set the column size property in the structure, for strings its the
        // number of characters. For fixed size columns its the characters
        // required to display it. For binary columns its the number of bytes.
        switch ( m_pColumns[ i ].sqlType )
        {
            default:
            case SQL_UNKNOWN_TYPE:      // 0

                throw SystemException( ERROR_INVALID_DATA, L"SQL_UNKNOWN_TYPE", __FUNCTION__ );

            // Fall through, fixed length columns, for display size
            case SQL_NUMERIC:           // 2
            case SQL_DECIMAL:           // 3

                m_pColumns[ i ].displayChars = ColumnSize;
                m_pColumns[ i ].cType        = SQL_C_NUMERIC;
                m_pColumns[ i ].sizeBytes    = sizeof ( SQL_NUMERIC_STRUCT );
                break;

            case SQL_INTEGER:           // 4

                m_pColumns[ i ].displayChars = ColumnSize;
                m_pColumns[ i ].cType        = SQL_C_SLONG;
                m_pColumns[ i ].sizeBytes    = sizeof ( SQLINTEGER );
                break;

            case SQL_SMALLINT:          // 5

                m_pColumns[ i ].displayChars = ColumnSize;
                m_pColumns[ i ].cType        = SQL_C_SSHORT;
                m_pColumns[ i ].sizeBytes    = sizeof ( SQLSMALLINT );
                break;

            case SQL_FLOAT:             // 6

                m_pColumns[ i ].displayChars = ColumnSize;
                m_pColumns[ i ].cType        = SQL_C_DOUBLE;
                m_pColumns[ i ].sizeBytes    = sizeof ( SQLFLOAT );
                break;

            case SQL_REAL:              // 7

                m_pColumns[ i ].displayChars = ColumnSize;
                m_pColumns[ i ].cType        = SQL_C_FLOAT;
                m_pColumns[ i ].sizeBytes    = sizeof ( SQLREAL );
                break;

            case SQL_DOUBLE:            // 8

                m_pColumns[ i ].displayChars = ColumnSize;
                m_pColumns[ i ].cType        = SQL_C_DOUBLE;
                m_pColumns[ i ].sizeBytes    = sizeof ( SQLDOUBLE );
                break;

            case SQL_TYPE_DATE:         // 91   ODBC 3.0

                m_pColumns[ i ].displayChars = ColumnSize;
                m_pColumns[ i ].cType        = SQL_C_TYPE_DATE;
                m_pColumns[ i ].sizeBytes    = sizeof ( SQL_DATE_STRUCT );
                break;

            case SQL_TYPE_TIME:         // 92

                m_pColumns[ i ].displayChars = ColumnSize;
                m_pColumns[ i ].cType        = SQL_C_TYPE_TIME;
                m_pColumns[ i ].sizeBytes    = sizeof ( SQL_TIME_STRUCT );
                break;

            case SQL_TYPE_TIMESTAMP:    // 93   ODBC 3.0

                m_pColumns[ i ].displayChars = ColumnSize;
                m_pColumns[ i ].cType        = SQL_C_TYPE_TIMESTAMP;
                m_pColumns[ i ].sizeBytes    = sizeof (SQL_TIMESTAMP_STRUCT);
                break;

            case SQL_BIGINT:            // -5

                m_pColumns[ i ].displayChars = ColumnSize;
                m_pColumns[ i ].cType        = SQL_C_SBIGINT;
                m_pColumns[ i ].sizeBytes    = sizeof ( SQLBIGINT );
                break;

            case SQL_TINYINT:           // -6

                m_pColumns[ i ].displayChars = ColumnSize;
                m_pColumns[ i ].cType        = SQL_C_STINYINT;
                m_pColumns[ i ].sizeBytes    = sizeof ( SQLSCHAR );
                break;

            case SQL_BIT:               // -7

                m_pColumns[ i ].displayChars = ColumnSize;
                m_pColumns[ i ].cType        = SQL_C_BIT;
                m_pColumns[ i ].sizeBytes    = sizeof ( SQLCHAR );
                break;

            case SQL_GUID:              // -11

                m_pColumns[ i ].displayChars = ColumnSize;
                m_pColumns[ i ].cType        = SQL_C_GUID;
                m_pColumns[ i ].sizeBytes    = sizeof ( SQLGUID );
                break;

            // String columns
            case SQL_CHAR:
            case SQL_VARCHAR:
            case SQL_LONGVARCHAR:

                m_pColumns[ i ].displayChars = ColumnSize;
                m_pColumns[ i ].cType        = SQL_C_CHAR;
                m_pColumns[ i ].sizeBytes    = ColumnSize;

                // Access 97 returns the Unicode character length for ANSI
                // strings. From v2000 onwards all text is Unicode, so if get
                // an ANSI data type scale by the size of a wchar_t
                if ( Dbms.IndexOfI( PXS_DBMS_NAME_ACCESS ) != PXS_MINUS_ONE )
                {
                    sizeBytes = m_pColumns[ i ].sizeBytes / sizeof ( wchar_t );
                    m_pColumns[ i ].sizeBytes = sizeBytes;
                }

                break;

            // Unicode strings
            case SQL_WCHAR:
            case SQL_WVARCHAR:
            case SQL_WLONGVARCHAR:

              sizeBytes = PXSMultiplySizeT( ColumnSize, sizeof ( SQLWCHAR ) );
              m_pColumns[ i ].displayChars = ColumnSize;
              m_pColumns[ i ].cType        = SQL_C_WCHAR;
              m_pColumns[ i ].sizeBytes    = sizeBytes;
              break;

            // For binary columns it the number of bytes. Use 1 byte because
            // do not want to display binary data but it an ODBC specification
            // error to have a zero-byte column.
            case SQL_BINARY:
            case SQL_VARBINARY:
            case SQL_LONGVARBINARY:
                m_pColumns[ i ].sizeBytes = 1;    // 1 byte
                m_pColumns[ i ].cType     = SQL_C_BINARY;
                break;
        }

        // Limit the size of data in a column
        PXSLimitSizeT( 1, MAX_COLUMN_LENGTH_BYTES, &m_pColumns[i].sizeBytes );

        ////////////////////////////////////////////////////////////////////////
        // Column properties using SQLColAttribute

        memset( szBuffer, 0, sizeof ( szBuffer ) );
        StringLength     = 0;
        NumericAttribute = 0;
        m_Odbc.ColAttribute( m_StatementHandle,
                             ColumnNumber,
                             SQL_DESC_TABLE_NAME,
                             szBuffer, ARRAYSIZE( szBuffer ), &StringLength, &NumericAttribute );
        szBuffer[ ARRAYSIZE( szBuffer ) - 1 ] = PXS_CHAR_NULL;
        StringCchCopy( m_pColumns[ i ].szTable,
                       ARRAYSIZE( m_pColumns[ i ].szTable ), szBuffer );

        // For decimal/numeric get the precision
        if ( ( m_pColumns[ i ].sqlType == SQL_NUMERIC ) ||
             ( m_pColumns[ i ].sqlType == SQL_DECIMAL )  )
        {
            NumericAttribute = 0;
            m_Odbc.ColAttribute( m_StatementHandle,
                                 ColumnNumber,
                                 SQL_DESC_PRECISION,
                                 nullptr, 0, nullptr, &NumericAttribute );    // SQLLEN
            // Verify bounds as precision is an unsigned char
            if ( ( NumericAttribute < 0 ) ||
                 ( NumericAttribute > UCHAR_MAX ) )     // UCHAR_MAX has no sign
            {
               throw BoundsException( L"NumericAttribute", __FUNCTION__ );
            }
            m_pColumns[ i ].bPrecision = PXSCastSqlLenToUInt8(NumericAttribute);
        }
    }
}

//===============================================================================================//
//  Description:
//      Get a pointer to the data value stored in the row cache at the
//      specified row and column
//
//  Parameters:
//      idxRow - zero based index of the column
//      idxColumn - zero based index of the column
//
//  Returns:
//      The pointer, NULL if the data is NULL
//===============================================================================================//
BYTE* OdbcRecordSet::GetDataPointerInCache( size_t idxRow, size_t idxColumn )
{
    size_t byteOffset = 0;

    // Must have a cache
    if ( m_pCache == nullptr )
    {
        throw FunctionException( L"m_pCache", __FUNCTION__ );
    }

    // Must have created some columns
    if ( m_pColumns == nullptr )
    {
        throw FunctionException( L"m_pColumns", __FUNCTION__ );
    }

    // Column bounds
    if ( idxColumn >= m_uNumColumns )
    {
        throw BoundsException( L"idxColumn", __FUNCTION__ );
    }

    // Row bounds, >= as want inclusive
    if ( ( idxRow <  m_idxCacheStart ) ||
         ( idxRow >= PXSAddSizeT( m_idxCacheStart, m_uNumRowsCached ) ) )
    {
        throw BoundsException( L"idxRow", __FUNCTION__ );
    }

    // Test if the row has been deleted
    if ( IsDeleted( idxRow ) )
    {
        return nullptr;
    }

    // Test for NULL
    if ( SQL_NULL_DATA == GetIndicatorValue( idxRow, idxColumn ) )
    {
        return nullptr;
    }

    // Get the location of the column's value in the buffer
    byteOffset = idxRow - m_idxCacheStart;      // Bounds were checked above
    byteOffset = PXSMultiplySizeT( byteOffset, m_uCacheRowSizeBytes );
    byteOffset = PXSAddSizeT( byteOffset, m_pColumns[ idxColumn ].rowOffsetBytes );

    return ( m_pCache + byteOffset );
}

//===============================================================================================//
//  Description:
//      Get the value of the indicator for the specified row and column
//
//  Parameters:
//      idxRow - zero based index of the column
//      idxColumn - zero based index of the column
//
//  Remarks:
//      The row must be in the cached data.
//
//  Returns:
//      The indicator's value
//===============================================================================================//
SQLLEN OdbcRecordSet::GetIndicatorValue( size_t idxRow, size_t idxColumn )
{
    size_t offset = 0, cacheSizeBytes = 0;
    SQLLEN StrLenOrInd = 0;

    // Check class scope variables, the columns array should have been created
    if ( ( m_pColumns == nullptr ) || ( m_uNumColumns == 0 ) )
    {
       throw FunctionException( L"m_pColumns/m_uNumColumns", __FUNCTION__ );
    }

    // Make sure the row is in bounds, inclusive test
    if ( ( idxRow <  m_idxCacheStart ) ||
         ( idxRow >= PXSAddSizeT( m_idxCacheStart, m_uNumRowsCached ) ) )
    {
        throw BoundsException( L"idxRow", __FUNCTION__ );
    }

    // Make sure the column is in bounds
    if ( idxColumn >= m_uNumColumns )
    {
        throw BoundsException( L"idxColumn", __FUNCTION__ );
    }

    // Get the location of column's indicator in the buffer
    offset = idxRow - m_idxCacheStart;    // Bounds were checked above
    offset = PXSMultiplySizeT( offset, m_uCacheRowSizeBytes );
    offset = PXSAddSizeT( offset, m_pColumns[ idxColumn ].indRowOffsetBytes );

    // Bounds check, the indicator is an SQLLEN and must be in the cache
    cacheSizeBytes = PXSMultiplySizeT( m_uCacheRowSizeBytes, m_uCacheSize );
    if ( PXSAddSizeT( offset, sizeof ( SQLLEN ) ) > cacheSizeBytes )
    {
        throw BoundsException( L"offset", __FUNCTION__ );
    }
    StrLenOrInd = *reinterpret_cast<SQLLEN*>( m_pCache + offset );

    return StrLenOrInd;
}

//===============================================================================================//
//  Description:
//      Clean up the array of field values
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void OdbcRecordSet::UnbindAndDeleteCache()
{
    size_t i = 0;

    // Order is important, first unbind
    if ( m_StatementHandle )
    {
        m_Odbc.FreeStmt( m_StatementHandle, SQL_UNBIND );
    }

    // Now delete the cache
    if ( m_pCache )
    {
        delete[] m_pCache;
        m_pCache = nullptr;       // Reset
    }

    // Reset the offsets
    if ( m_pColumns )
    {
        for ( i = 0; i < m_uNumColumns; i++ )
        {
            m_pColumns[ i ].rowOffsetBytes    = 0;
            m_pColumns[ i ].indRowOffsetBytes = 0;
        }
    }

    // Reset the count of rows cached and the size of the row
    m_uNumRowsCached     = 0;
    m_uCacheRowSizeBytes = 0;
}

//===============================================================================================//
//  Description:
//      Reset the cached field values array to zero
//
//  Parameters:
//      idxOffset - index of first record in array to be zeroed
//
//  Remarks:
//      Not deleting the cache, just re-setting to zero from
//      idxOffset to the end of the array
//
//  Returns:
//      void
//===============================================================================================//
void OdbcRecordSet::ZeroCache( size_t idxOffset )
{
    size_t numBytes = 0, byteOffset = 0;

    if ( ( m_pCache == nullptr ) || ( m_uCacheRowSizeBytes == 0 ) )
    {
        return;     // Nothing to do
    }

    // Bounds check
    if ( idxOffset >= m_uCacheSize )
    {
        throw BoundsException( L"idxOffset", __FUNCTION__ );
    }
    byteOffset = PXSMultiplySizeT( idxOffset, m_uCacheRowSizeBytes );
    numBytes   = PXSMultiplySizeT( m_uCacheSize - idxOffset, m_uCacheRowSizeBytes );
    memset( m_pCache + byteOffset, 0, numBytes );
}
