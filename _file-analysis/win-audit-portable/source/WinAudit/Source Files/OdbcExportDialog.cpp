///////////////////////////////////////////////////////////////////////////////////////////////////
//
// ODBC Database Export Class Implementation
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
#include "WinAudit/Header Files/OdbcExportDialog.h"

// 2. C System Files
#include <limits.h>
#include <sqlext.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/WaitCursor.h"

// 5. This Project
#include "WinAudit/Header Files/AuditDatabase.h"
#include "WinAudit/Header Files/DatabaseAdministrationDialog.h"
#include "WinAudit/Header Files/OdbcRecordSet.h"
#include "WinAudit/Header Files/Resources.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default Constructor
OdbcExportDialog::OdbcExportDialog()
                 :m_bDidExport(),
                  m_idxProgressMessageStatic( PXS_MINUS_ONE ),
                  m_uNumColumns( 0 ),
                  m_pColumnProps( nullptr ),
                  m_pRowStatus( nullptr ),
                  m_AuditMasterRecord(),
                  m_ComputerMasterRecord(),
                  m_AuditRecords(),
                  m_Settings(),
                  m_MaxErrorRateSpinner(),
                  m_MaxAffectedRowsSpinner(),
                  m_ConnectTimeoutSpinner(),
                  m_QueryTimeoutSpinner(),
                  m_CloseButton(),
                  m_BrowseButton(),
                  m_AdminButton(),
                  m_ExportButton(),
                  m_DatabaseNameTextField(),
                  m_MySQLComboBox(),
                  m_PostgreComboBox(),
                  m_AccessOption(),
                  m_SqlServerOption(),
                  m_MySQLOption(),
                  m_PostgreSQLOption(),
                  m_ProgressBar()
{
    try
    {
        SetSize( 500, 410 );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy Constructor - not allowed, no implementation

// Destructor
OdbcExportDialog::~OdbcExportDialog()
{
    // Clean up
    FreeAuditDataBindBuffers();
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
//      Get if records were exported to the database
//
//  Parameters:
//      none
//
//  Returns:
//      true of records exported, otherwise false
//===============================================================================================//
bool OdbcExportDialog::DidExport() const
{
    return m_bDidExport;
}

//===============================================================================================//
//  Description:
//      Export the records to the database using a full connection string.
//
//  Parameters:
//      ConnectionString - the connection string
//
//  Remarks:
//      Does not prompt the user
//
//  Returns:
//        void
//===============================================================================================//
void OdbcExportDialog::DoExport( const String& ConnectionString )
{
    String ResultMsg;
    AuditDatabase Database;

    Database.Connect( ConnectionString,
                      m_Settings.connectTimeoutSecs,
                      m_Settings.queryTimeoutSecs,
                      nullptr );        // without prompting

    ExportRecordsToDatabase( &Database, &ResultMsg );
    PXSLogAppInfo1( L"Database export result: %%1", ResultMsg );
}

//===============================================================================================//
//  Description:
//      Export the records to the database
//
//  Parameters:
//      pDatabase      - the database object
//      pResultMessage - receives a text message
//
//  Remarks:
//      Audit record status:
//      0 = inserting
//      1 = inserted
//      2 = deleting
//      3 = active
//
//  Returns:
//      void
//===============================================================================================//
void OdbcExportDialog::ExportRecordsToDatabase( AuditDatabase* pDatabase, String* pResultMessage )
{
    Odbc      OdbcObject;
    DWORD     startTickCount = 0, errorRate = 0, numRowsAdded = 0, numErrors =0;
    size_t    i = 0, numAuditRecords = 0, temp = 0;
    String    ErrorMessage, Value, SqlCleanValue, SqlQuery;
    String    ComputerIDString, AuditIDString, RecordString, ErrorsCount;
    String    MilliSecs, AddedCount;
    Formatter Format;
    AuditRecord  Record;
    StringArray  Values;
    HSTMT        hStmt  = nullptr;
    SQLHDBC      hDBC   = nullptr;
    SQLINTEGER   auditID = 0, computerID = 0;
    SQLUSMALLINT columnNumber = 0;

    if ( pDatabase == nullptr )
    {
        throw ParameterException( L"pDatabase", __FUNCTION__ );
    }

    if ( pResultMessage == nullptr )
    {
        throw ParameterException( L"pResultMessage", __FUNCTION__ );
    }
    *pResultMessage = PXS_STRING_EMPTY;

    numAuditRecords = m_AuditRecords.GetSize();
    if ( numAuditRecords == 0 )
    {
        throw FunctionException( L"m_AuditRecords", __FUNCTION__ );
    }

    Value = Format.SizeT( numAuditRecords );
    PXSLogAppInfo1( L"Inserting %%1 audit records.", Value );
    if ( numAuditRecords > m_Settings.maxAffectedRows )
    {
        PXSGetResourceString( PXS_IDS_1237_TOO_MANY_ROWS_AFFECTED, &ErrorMessage );
        throw SystemException( ERROR_INVALID_DATA, ErrorMessage.c_str(), __FUNCTION__ );
    }

    if ( pDatabase->IsConnected() == false )
    {
        PXSGetResourceString( PXS_IDS_1233_DB_NOT_CONNECTED, &ErrorMessage );
        throw SystemException( ERROR_INVALID_FUNCTION, ErrorMessage.c_str(), __FUNCTION__ );
    }

    if ( PXS_CATEGORY_UKNOWN == m_AuditMasterRecord.GetCategoryID() )
    {
        throw FunctionException( L"m_AuditMasterRecord", __FUNCTION__ );
    }
    m_AuditMasterRecord.ToString( &Value );
    PXSLogAppInfo1( L"Audit Master Record: '%%1'", Value );
    m_ProgressBar.SetPercentage( 30 );

    if ( PXS_CATEGORY_UKNOWN == m_ComputerMasterRecord.GetCategoryID() )
    {
        throw FunctionException( L"m_ComputerMasterRecord", __FUNCTION__ );
    }
    m_ComputerMasterRecord.ToString( &Value );
    PXSLogAppInfo1( L"Computer Master Record: '%%1'", Value );
    m_ProgressBar.SetPercentage( 40 );

    // Start database work
    startTickCount = GetTickCount();
    FreeAuditDataBindBuffers();
    AllocateAuditDataBindBuffers( pDatabase );
    pDatabase->BeginTrans();

    // Identify the Computer_ID, if none then create one
    computerID = pDatabase->IdentifyComputerID( m_ComputerMasterRecord );
    if ( computerID == 0 )
    {
        computerID = pDatabase->InsertComputerMaster( m_ComputerMasterRecord );
    }
    ComputerIDString = Format.Int32( computerID );
    PXSLogAppInfo1( L"The Computer_ID for this computer is: %%1.",
                    ComputerIDString );
    m_ProgressBar.SetPercentage( 60 );

    // Get the Audit ID
    auditID = pDatabase->InsertAuditMaster( computerID, m_AuditMasterRecord );
    AuditIDString = Format.Int32( auditID );
    PXSLogAppInfo1( L"The Audit_ID for this audit is: %%1.", AuditIDString );
    m_ProgressBar.SetPercentage( 70 );

    // Allocate a statement and set its properties
    hDBC = pDatabase->GetConnectionHandle();
    OdbcObject.AllocHandle( SQL_HANDLE_STMT, hDBC, &hStmt );
    try
    {
        // Optional query time out. Not supported by all databases
        if ( pDatabase->SupportsQueryTimeOut() )
        {
            OdbcObject.SetStmtAttr(
               hStmt,
               SQL_ATTR_QUERY_TIMEOUT,
               (SQLPOINTER)(DWORD_PTR)m_Settings.queryTimeoutSecs,  // TYPE CAST
               SQL_IS_UINTEGER );
        }
        OdbcObject.SetStmtAttr(
                            hStmt,
                            SQL_ATTR_ROW_BIND_TYPE,
                            nullptr,            // SQL_BIND_BY_COLUMN = 0
                            SQL_IS_UINTEGER );
        OdbcObject.SetStmtAttr(
                            hStmt,
                            SQL_ATTR_ROW_ARRAY_SIZE,
                            reinterpret_cast<SQLPOINTER>( numAuditRecords ),
                            SQL_IS_UINTEGER );
        OdbcObject.SetStmtAttr(
                            hStmt,
                            SQL_ATTR_ROW_STATUS_PTR,
                            reinterpret_cast<SQLPOINTER>( m_pRowStatus ),
                            SQL_IS_POINTER );
        OdbcObject.SetStmtAttr(
                            hStmt,
                            SQL_ATTR_CONCURRENCY,
                            reinterpret_cast<SQLPOINTER>( SQL_CONCUR_ROWVER ),
                            SQL_IS_UINTEGER );
        OdbcObject.SetStmtAttr(
                        hStmt,
                        SQL_ATTR_CURSOR_TYPE,
                        reinterpret_cast<SQLPOINTER>(SQL_CURSOR_KEYSET_DRIVEN),
                        SQL_IS_UINTEGER );

        // Return an empty result set
        SqlQuery = L"SELECT * FROM Audit_Data WHERE 1 = 2";
        OdbcObject.ExecDirect( hStmt, SqlQuery );
        OdbcObject.Fetch( hStmt );
        m_ProgressBar.SetPercentage( 80 );

        // Bind and add
        FillAuditDataBindBuffers( auditID, computerID );
        for ( i = 0; i < m_uNumColumns; i++ )
        {
            columnNumber = PXSCastSizeTToUInt16( i );   // SQLUSMALLINT = USHORT
            columnNumber = PXSAddUInt16( columnNumber, 1 );
            OdbcObject.BindCol( hStmt,
                                columnNumber,
                                m_pColumnProps[ i ].TargetType,
                                m_pColumnProps[ i ].TargetValuePtr,
                                m_pColumnProps[ i ].BufferLength,
                                m_pColumnProps[ i ].StrLen_or_Ind );
        }
        OdbcObject.BulkOperations( hStmt, SQL_ADD );
        m_ProgressBar.SetPercentage( 90 );
    }
    catch ( const Exception& )
    {
        OdbcObject.FreeHandle( SQL_HANDLE_STMT, hStmt );
        pDatabase->RollbackTrans();
        throw;
    }
    OdbcObject.FreeHandle( SQL_HANDLE_STMT, hStmt );

    // Count the errors
    for ( i = 0; i < numAuditRecords; i++ )
    {
        if ( m_pRowStatus[ i ] == SQL_ROW_ADDED )
        {
            numRowsAdded = PXSAddUInt32( numRowsAdded, 1 );
        }
        else
        {
            numErrors = PXSAddUInt32( numErrors, 1 );
            Record    = m_AuditRecords.Get( i );
            Record.ToString( &RecordString );
            PXSLogAppWarn1( L"Did not insert record '%%1'", RecordString );
        }
    }

    // Test for error
    temp      = PXSMultiplySizeT( 100, numErrors ) / numAuditRecords;
    errorRate = PXSCastSizeTToUInt32( temp );
    if ( errorRate <= m_Settings.maxErrorRate )
    {
        pDatabase->CommitTrans();

        // Update the computer_master table with the audit id.
        SqlQuery  = L"UPDATE Computer_Master SET Last_Audit_ID=";
        SqlQuery += AuditIDString;
        SqlQuery += L" WHERE Computer_ID=";
        SqlQuery += ComputerIDString;
        pDatabase->BeginTrans();
        pDatabase->ExecuteTrans( SqlQuery, 1 );   // Limit to 1 row
        pDatabase->CommitTrans();

        // Result message
        AddedCount  = Format.UInt32( numRowsAdded );
        ErrorsCount = Format.UInt32( numErrors );
        MilliSecs   = Format.UInt32( GetTickCount() - startTickCount );
        if ( g_pApplication )
        {
            g_pApplication->GetResourceString3( PXS_IDS_1236_INSERT_RESULT_MESSAGE,
                                                AddedCount,
                                                ErrorsCount, MilliSecs, pResultMessage );
        }
        PXSLogAppInfo( pResultMessage->c_str() );
        m_bDidExport = true;
        m_ExportButton.SetEnabled( false );
    }
    else
    {
        pDatabase->RollbackTrans();
        PXSGetResourceString( PXS_IDS_1235_TOO_MANY_ERRORS, &ErrorMessage );
        throw SystemException( ERROR_INVALID_DATA, ErrorMessage.c_str(), __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Get the settings used by the dialog
//
//  Parameters:
//      pSettings - receives the configuration settings
//
//  Returns:
//      void
//===============================================================================================//
void OdbcExportDialog::GetConfigurationSettings(
                                       ConfigurationSettings* pSettings ) const
{
    if ( pSettings == nullptr )
    {
        throw ParameterException( L"pSettings", __FUNCTION__ );
    }
    *pSettings = m_Settings;
}

//===============================================================================================//
//  Description:
//      Set the audit records to send to the database
//
//  Parameters:
//      AuditMasterRecord    - the record for the audit_master table
//      ComputerMasterRecord - the record for the computer_master table
//      AuditRecords         - array of audit records
//
//  Returns:
//        void
//===============================================================================================//
void OdbcExportDialog::SetAuditRecords( const AuditRecord& AuditMasterRecord,
                                        const AuditRecord& ComputerMasterRecord,
                                        const TArray< AuditRecord >& AuditRecords )
{
     m_AuditMasterRecord    = AuditMasterRecord;
     m_ComputerMasterRecord = ComputerMasterRecord;
     m_AuditRecords.RemoveAll();
     m_AuditRecords.Append( AuditRecords );
}

//===============================================================================================//
//  Description:
//      Set the settings used by this dialog
//
//  Parameters:
//      Settings - the configuration settings
//
//  Returns:
//      void
//===============================================================================================//
void OdbcExportDialog::SetConfigurationSettings( const ConfigurationSettings& Settings )
{
    m_Settings = Settings;

    // Tidy up those settings that pertain to this dialog
    if ( ( m_Settings.connectTimeoutSecs < PXS_DB_CONNECT_TIMEOUT_SECS_MIN ) ||
         ( m_Settings.connectTimeoutSecs > PXS_DB_CONNECT_TIMEOUT_SECS_MAX )  )
    {
        m_Settings.connectTimeoutSecs = PXS_DB_CONNECT_TIMEOUT_SECS_DEF;
    }

    if ( ( m_Settings.maxAffectedRows < PXS_DB_MAX_AFFECTED_ROWS_MIN ) ||
         ( m_Settings.maxAffectedRows > PXS_DB_MAX_AFFECTED_ROWS_MAX )  )
    {
        m_Settings.maxAffectedRows = PXS_DB_MAX_AFFECTED_ROWS_DEFAULT;
    }

    if ( m_Settings.maxErrorRate > PXS_DB_MAX_ERROR_RATE_MAX )
    {
        m_Settings.maxErrorRate = PXS_DB_MAX_ERROR_RATE_DEFAULT;
    }

    if ( ( m_Settings.queryTimeoutSecs < PXS_DB_QUERY_TIMEOUT_SECS_MIN ) ||
         ( m_Settings.queryTimeoutSecs > PXS_DB_QUERY_TIMEOUT_SECS_MAX )  )
    {
        m_Settings.queryTimeoutSecs = PXS_DB_QUERY_TIMEOUT_SECS_DEF;
    }

    m_Settings.DatabaseName.Trim();
    m_Settings.DBMS.Trim();
    m_Settings.LastReportName.Trim();
    m_Settings.MySqlDriver.Trim();
    m_Settings.PostgreSqlDriver.Trim();
    m_Settings.ServerName.Trim();
    m_Settings.UID.Trim();
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Handle WM_COMMAND events.
//
//  Parameters:
//      wParam -
//      lParam -
//
//  Returns:
//      0 if handled, else non-zero.
//===============================================================================================//
LRESULT OdbcExportDialog::CommandEvent( WPARAM wParam, LPARAM lParam )
{
    WORD    wID     = LOWORD(wParam);  // Item, control, or accelerator
    HWND    hWnd    = reinterpret_cast<HWND>( lParam );    // Window handle
    LRESULT result  = 0;
    DWORD   filterIndex = 0;
    String  DatabaseName;

    // Ensure controls have been created
    if ( m_bControlsCreated == false ) return result;

    if ( IDCANCEL == wID )
    {
        EndDialog( m_hWindow, IDCANCEL );
    }
    else if ( IsClickFromButton( m_CloseButton, wParam, lParam ) )
    {
        UpdateConfigurationSettings();
        EndDialog( m_hWindow, IDOK );
    }
    else if ( ( hWnd == m_AccessOption.GetHwnd()     ) ||
              ( hWnd == m_SqlServerOption.GetHwnd()  ) ||
              ( hWnd == m_MySQLOption.GetHwnd()      ) ||
              ( hWnd == m_PostgreSQLOption.GetHwnd() )  )
    {
        m_BrowseButton.SetEnabled( m_AccessOption.GetState() );
        m_MySQLComboBox.SetEnabled( m_MySQLOption.GetState() );
        m_PostgreComboBox.SetEnabled( m_PostgreSQLOption.GetState() );
    }
    else if ( IsClickFromButton( m_BrowseButton, wParam, lParam ) )
    {
        filterIndex = 3;      // *.mdb
        if ( PXSPromptForFilePath( m_hWindow,
                                   true ,
                                   false,
                                   L"All Files (*.*)\0*.*\0"
                                   L"Microsoft Access (*.accdb)\0*.accdb\0"
                                   L"Microsoft Access (*.mdb)\0*.mdb\0",
                                   &filterIndex, nullptr, L"* accdb mdb", &DatabaseName ) )
        {
            DatabaseName.Trim();
            m_DatabaseNameTextField.SetText( DatabaseName );
        }
    }
    else if ( IsClickFromButton( m_AdminButton, wParam, lParam ) )
    {
        ShowAdminDialog();
    }
    else if ( IsClickFromButton( m_ExportButton, wParam, lParam ) )
    {
        ExportRecords();
    }
    else
    {
        result = 1;     //  Not handled, return non-zero
    }

    return result;
}

//===============================================================================================//
//  Description:
//      Handle the WM_INITDIALOG message
//
//  Parameters:
//      None
//
//  Returns:
//        void.
//===============================================================================================//
void OdbcExportDialog::InitDialogEvent()
{
    int     CONTROL_HEIGHT = 20, LINE_HEIGHT = 23, GAP = 10;
    bool    rtlReading;
    Font    FontObject;
    Odbc    OdbcObject;
    SIZE    buttonSize = { 0, 0 };
    SIZE    dialogSize = { 0, 0 };
    RECT    bounds     = { 0, 0, 0, 0 };
    POINT   location   = { 0, 0 };
    DWORD   i = 0, frameShape = PXS_SHAPE_FRAME;
    size_t  numDrivers = 0;
    String  Text, TempString;
    StringArray   OdbcDrivers;
    StaticControl Static;

    if ( m_hWindow == nullptr ) return;

    // Properties
    buttonSize.cx = 80;
    buttonSize.cy = 25;
    GetClientSize( &dialogSize);
    rtlReading  = IsRightToLeftReading();
    if ( rtlReading )
    {
        frameShape = PXS_SHAPE_FRAME_RTL;
    }

    try
    {
        OdbcObject.GetInstalledDrivers( &OdbcDrivers );
    }
    catch ( const Exception& e )
    {
        PXSShowExceptionDialog( e, m_hWindow );
    }

    // Title Caption, centred
    bounds.left   = GAP;
    bounds.right  = dialogSize.cx - GAP;
    bounds.top    = 5;
    bounds.bottom = 40;
    Static.SetBounds( bounds );
    FontObject.SetPointSize( 12, nullptr );
    FontObject.SetBold( true );
    FontObject.SetUnderlined( true );
    FontObject.Create();
    Static.SetFont( FontObject );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    PXSGetResourceString( PXS_IDS_1080_EXPORT_AUDIT_TO_DB, &Text );
    Static.SetText( Text );
    m_Statics.Add( Static );
    Static.Reset();

    // Frame for Database management system
    PXSGetResourceString( PXS_IDS_1081_DATABASE, &Text );
    bounds.left   = GAP;
    bounds.top    = bounds.bottom;
    bounds.bottom = bounds.top + ( 7 * LINE_HEIGHT );
    Static.SetBounds( bounds );
    Static.SetShape( frameShape );
    Static.SetText( Text );
    m_Statics.Add( Static );
    Static.Reset();

    // Access
    Text = PXS_DBMS_NAME_ACCESS;
    bounds.left  += GAP;
    bounds.top   += GAP + GAP;
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    m_AccessOption.StartNewGroup();     // Must set style before creation
    m_AccessOption.SetBounds( bounds.left,
                              bounds.top, dialogSize.cx / 2 - GAP - GAP, CONTROL_HEIGHT );
    m_AccessOption.Create( m_hWindow );
    m_AccessOption.SetText( Text );

    // SQL Server
    Text = PXS_DBMS_NAME_SQL_SERVER;
    bounds.top  += LINE_HEIGHT;
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    m_SqlServerOption.SetBounds( bounds.left,
                                 bounds.top, dialogSize.cx / 2 - GAP - GAP, CONTROL_HEIGHT );
    m_SqlServerOption.Create( m_hWindow );
    m_SqlServerOption.SetText( Text );

    // MySQL
    Text = PXS_DBMS_NAME_MYSQL;
    bounds.top += LINE_HEIGHT;
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    m_MySQLOption.SetBounds( bounds.left,
                             bounds.top, dialogSize.cx / 2 - GAP - GAP, CONTROL_HEIGHT );
    m_MySQLOption.Create( m_hWindow );
    m_MySQLOption.SetText( Text );

    // MySQL ODBC driver combo list
    m_MySQLComboBox.SetBounds( dialogSize.cx / 2  + GAP,
                               bounds.top, dialogSize.cx / 2 - GAP - GAP - GAP, 100 );
    m_MySQLComboBox.Create( m_hWindow );

    // Populate the list of MySQL drop down list
    numDrivers = OdbcDrivers.GetSize();
    for ( i = 0; i < numDrivers; i++ )
    {
        TempString = OdbcDrivers.Get( i );
        if ( TempString.IndexOfI( PXS_DBMS_NAME_MYSQL ) != PXS_MINUS_ONE )
        {
            m_MySQLComboBox.Add( TempString.c_str() );
        }
    }

    // PostgreSQL
    Text = PXS_DBMS_NAME_POSTGRE_SQL;
    bounds.top += LINE_HEIGHT;
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    m_PostgreSQLOption.SetBounds( bounds.left,
                                  bounds.top, dialogSize.cx / 2 - GAP - GAP, CONTROL_HEIGHT );
    m_PostgreSQLOption.Create( m_hWindow );
    m_PostgreSQLOption.SetText( Text );

    // PostgreSQL ODBC driver combo list
    m_PostgreComboBox.SetBounds( dialogSize.cx / 2  + GAP,
                                 bounds.top,
                                 dialogSize.cx / 2 - GAP - GAP - GAP, 100 );
    m_PostgreComboBox.Create( m_hWindow );
    for ( i = 0; i < OdbcDrivers.GetSize(); i++ )
    {
        TempString = OdbcDrivers.Get( i );
        if ( TempString.IndexOfI( PXS_DBMS_NAME_POSTGRE_SQL ) != PXS_MINUS_ONE )
        {
            m_PostgreComboBox.Add( TempString.c_str() );
        }
    }

    // Database name label
    PXSGetResourceString( PXS_IDS_1082_DATABASE_NAME_COLON, &Text );
    bounds.top += LINE_HEIGHT;
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    Static.SetBounds( bounds );
    Static.SetText( Text );
    m_Statics.Add( Static );
    Static.Reset();

    // Database text box
    bounds.top    = bounds.bottom;
    bounds.right  = dialogSize.cx - GAP - GAP - buttonSize.cx - GAP;
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    m_DatabaseNameTextField.Create( m_hWindow );
    m_DatabaseNameTextField.SetBounds( bounds );
    m_DatabaseNameTextField.SetText( m_Settings.DatabaseName );
    m_DatabaseNameTextField.SetTextMaxLength( MAX_PATH );

    // Add the button to browse for file based databases, e.g. Access
    PXSGetResourceString( PXS_IDS_1083_BROWSE, &Text );
    m_BrowseButton.Create( m_hWindow );
    m_BrowseButton.SetLocation( bounds.right + GAP, bounds.top - 4 );
    m_BrowseButton.SetText( Text );

    // Frame for export options
    PXSGetResourceString( PXS_IDS_1084_EXPORT_DATA_OPTIONS, &Text );
    bounds.left   = GAP;
    bounds.right  = dialogSize.cx - GAP;
    bounds.top    = bounds.bottom + LINE_HEIGHT;
    bounds.bottom = bounds.top + ( 2 * LINE_HEIGHT );
    Static.SetBounds( bounds );
    Static.SetShape( frameShape );
    Static.SetText( Text );
    m_Statics.Add( Static );
    Static.Reset();

    // Maximum error rate
    PXSGetResourceString( PXS_IDS_1085_MAXIMUM_ERROS_PERCENT, &Text );
    bounds.left  += GAP;
    bounds.top   += GAP + GAP;
    bounds.right  = dialogSize.cx / 2;
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    Static.SetBounds( bounds );
    Static.SetSingleLineText( Text );
    Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    m_Statics.Add( Static );
    Static.Reset();

    m_MaxErrorRateSpinner.SetBounds( ( dialogSize.cx / 2 ) - 65,
                                       bounds.top, 65, CONTROL_HEIGHT );
    m_MaxErrorRateSpinner.Create( m_hWindow );
    m_MaxErrorRateSpinner.SetRange( PXS_DB_MAX_ERROR_RATE_MIN, PXS_DB_MAX_ERROR_RATE_MAX );
    m_MaxErrorRateSpinner.SetValue( m_Settings.maxErrorRate );

    // Maximum allowed affected records
    PXSGetResourceString( PXS_IDS_1086_MAX_AFFECTED_ROWS, &Text );
    bounds.left   = dialogSize.cx / 2 + GAP;
    bounds.right  = dialogSize.cx - GAP;
    Static.SetBounds( bounds );
    Static.SetSingleLineText( Text );
    Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    m_Statics.Add( Static );
    Static.Reset();

    m_MaxAffectedRowsSpinner.SetBounds( dialogSize.cx - GAP - GAP - 65,
                                         bounds.top, 65, CONTROL_HEIGHT );
    m_MaxAffectedRowsSpinner.Create( m_hWindow );
    m_MaxAffectedRowsSpinner.SetRange( PXS_DB_MAX_AFFECTED_ROWS_MIN, PXS_DB_MAX_AFFECTED_ROWS_MAX );
    m_MaxAffectedRowsSpinner.SetValue( m_Settings.maxAffectedRows );

    // Frame for Connection Options
    PXSGetResourceString( PXS_IDS_1087_CONNECTION_OPTIONS, &Text );
    bounds.left   = GAP;
    bounds.right  = dialogSize.cx - GAP;
    bounds.top    = bounds.bottom + LINE_HEIGHT;
    bounds.bottom = bounds.top + ( 2 * LINE_HEIGHT );
    Static.SetBounds( bounds );
    Static.SetShape( frameShape );
    Static.SetText( Text );
    m_Statics.Add( Static );
    Static.Reset();

    // Database login timeout
    PXSGetResourceString( PXS_IDS_1088_CONNECTION_TIMEOUT_SECS, &Text );
    bounds.left  += GAP;
    bounds.top   += GAP + GAP;
    bounds.right  = dialogSize.cx / 2;
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    Static.SetBounds( bounds );
    Static.SetSingleLineText( Text );
    Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    m_Statics.Add( Static );
    Static.Reset();

    m_ConnectTimeoutSpinner.SetBounds( ( dialogSize.cx / 2 ) - 65,
                                         bounds.top, 65, CONTROL_HEIGHT );
    m_ConnectTimeoutSpinner.Create( m_hWindow );
    m_ConnectTimeoutSpinner.SetRange( PXS_DB_CONNECT_TIMEOUT_SECS_MIN,
                                      PXS_DB_CONNECT_TIMEOUT_SECS_MAX );
    m_ConnectTimeoutSpinner.SetValue( m_Settings.connectTimeoutSecs );

    // Query timeout
    PXSGetResourceString( PXS_IDS_1089_QUERY_TIMEOUT_SECS, &Text );
    bounds.left   = dialogSize.cx / 2 + GAP;
    bounds.right  = dialogSize.cx - GAP;
    Static.SetBounds( bounds );
    Static.SetSingleLineText( Text );
    Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    m_Statics.Add( Static );
    Static.Reset();

    m_QueryTimeoutSpinner.SetBounds( dialogSize.cx - GAP - GAP - 65,
                                     bounds.top, 65, CONTROL_HEIGHT );
    m_QueryTimeoutSpinner.Create( m_hWindow );
    m_QueryTimeoutSpinner.SetRange( PXS_DB_QUERY_TIMEOUT_SECS_MIN, PXS_DB_QUERY_TIMEOUT_SECS_MAX );
    m_QueryTimeoutSpinner.SetValue( m_Settings.queryTimeoutSecs );

    // Progress bar
    bounds.left   = GAP;
    bounds.top    = dialogSize.cy - GAP - buttonSize.cy - GAP - CONTROL_HEIGHT - GAP;
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    bounds.right  = dialogSize.cx - GAP;
    m_ProgressBar.Create( m_hWindow );
    m_ProgressBar.SetVisible( false );
    m_ProgressBar.SetBounds( bounds );
    m_ProgressBar.SetForeground( PXS_COLOUR_GREY );

    // Invisible label in the same location as the progress bar
    m_idxProgressMessageStatic = m_Statics.GetSize();
    Static.SetBounds( bounds );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    m_Statics.Add( Static );
    Static.Reset();

    // Horizontal line
    bounds.left   = GAP;
    bounds.right  = dialogSize.cx - GAP;
    bounds.top    = dialogSize.cy - GAP - buttonSize.cy - GAP;
    bounds.bottom = bounds.top + 2;
    Static.SetBounds( bounds );
    Static.SetShape( PXS_SHAPE_SUNK );
    m_Statics.Add( Static );
    Static.Reset();

    ////////////////////////////////////////////
    // Buttons

    location.y = dialogSize.cy - 10 - buttonSize.cy;

    // Horizontal line
    bounds.left   = 10;
    bounds.right  = dialogSize.cx - 10;
    bounds.top    = location.y - 10;
    bounds.bottom = bounds.top + 2;
    Static.SetBounds( bounds );
    Static.SetShape( PXS_SHAPE_SUNK );
    m_Statics.Add( Static );
    Static.Reset();

    // Button to administer database
    PXSGetResourceString( PXS_IDS_1090_ADMINISTRATION, &Text );
    location.x = dialogSize.cx - ( 3 * ( GAP + buttonSize.cx ) );
    m_AdminButton.SetLocation( location );
    m_AdminButton.Create( m_hWindow );
    m_AdminButton.SetText( Text );

    // Button to export records to the database
    PXSGetResourceString( PXS_IDS_1091_EXPORT, &Text );
    location.x += ( GAP + buttonSize.cx );
    m_ExportButton.SetLocation( location );
    m_ExportButton.Create( m_hWindow );
    m_ExportButton.SetText( Text );

    // Close button
    location.x += ( GAP + buttonSize.cx );
    m_CloseButton.SetLocation( location );
    m_CloseButton.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_130_CLOSE, &Text );
    m_CloseButton.SetText( Text );

    // Get here, so say controls created
    m_bControlsCreated = true;

    // Set controls
    if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_MYSQL ) == 0 )
    {
        m_MySQLOption.SetState( true );
    }
    else if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_POSTGRE_SQL ) == 0 )
    {
        m_PostgreSQLOption.SetState( true );
    }
    else if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_SQL_SERVER ) == 0 )
    {
        m_SqlServerOption.SetState( true );
    }
    else
    {
        // Default
        m_AccessOption.SetState( true );
    }
    m_BrowseButton.SetEnabled( m_AccessOption.GetState() );
    m_MySQLComboBox.SetSelectedString( m_Settings.MySqlDriver );
    m_PostgreComboBox.SetSelectedString( m_Settings.PostgreSqlDriver );
    m_MySQLComboBox.SetEnabled( m_MySQLOption.GetState() );
    m_PostgreComboBox.SetEnabled( m_PostgreSQLOption.GetState() );

    if ( m_AuditRecords.GetSize() == 0 )
    {
        m_ExportButton.SetEnabled( false );     // Nothing to do
    }

    // Mirror for RTL
    if ( rtlReading )
    {
        RtlStatics();
        m_MaxErrorRateSpinner.RtlMirror( dialogSize.cx );
        m_MaxAffectedRowsSpinner.RtlMirror( dialogSize.cx );
        m_ConnectTimeoutSpinner.RtlMirror( dialogSize.cx );
        m_QueryTimeoutSpinner.RtlMirror( dialogSize.cx );
        m_CloseButton.RtlMirror( dialogSize.cx );
        m_BrowseButton.RtlMirror( dialogSize.cx );
        m_AdminButton.RtlMirror( dialogSize.cx );
        m_ExportButton.RtlMirror( dialogSize.cx );
        m_DatabaseNameTextField.RtlMirror( dialogSize.cx );
        m_MySQLComboBox.RtlMirror( dialogSize.cx );
        m_PostgreComboBox.RtlMirror( dialogSize.cx );
        m_AccessOption.RtlMirror( dialogSize.cx );
        m_SqlServerOption.RtlMirror( dialogSize.cx );
        m_MySQLOption.RtlMirror( dialogSize.cx );
        m_PostgreSQLOption.RtlMirror( dialogSize.cx );
        m_ProgressBar.RtlMirror( dialogSize.cx );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Allocate class scope buffers to bind to the Audit_Data table
//
//  Parameters:
//      pDatabase - the audit database
//
//  Returns:
//      void
//===============================================================================================//
void OdbcExportDialog::AllocateAuditDataBindBuffers( AuditDatabase* pDatabase )
{
    const size_t  NUM_COLUMNS = 54;     // 4 integer + 50 string columns
    Odbc          OdbcObject;
    size_t        i = 0, length = 0, numBytes = 0, numRecords = 0, columnCount;
    String        SqlQuery, ErrorMessage;
    Formatter     Format;
    SQLSMALLINT   columnType = 0;
    OdbcRecordSet RecordSet;

    if ( pDatabase == nullptr )
    {
        throw ParameterException( L"pDatabase", __FUNCTION__ );
    }
    // Get the Audit_Data table's meta data
    SqlQuery  = L"SELECT * FROM v_Audit_Data_Empty";
    RecordSet.Open( SqlQuery, pDatabase );
    columnCount = RecordSet.GetColumnCount();

    // There should be at least 4 integer columns followed by 50 string columns
    if ( columnCount < NUM_COLUMNS )
    {
        throw SystemException( ERROR_DATABASE_FAILURE,
                               L"m_uNumColumns < 54", __FUNCTION__ );
    }

    for ( i = 0; i < NUM_COLUMNS; i++ )
    {
        columnType = RecordSet.GetColumnSqlType( i );
        if ( i < 4 )
        {
            if ( columnType != SQL_INTEGER )
            {
                throw SystemException( ERROR_INVALID_DATATYPE, L"SQLINTEGER", __FUNCTION__ );
            }
        }
        else
        {
            if ( OdbcObject.IsSqlStringType( columnType ) == false )
            {
                throw SystemException( ERROR_INVALID_DATATYPE,
                                       L"Odbc::IsStringType", __FUNCTION__ );
            }
        }
    }

    // Free
    FreeAuditDataBindBuffers();

    // Alloc the columns
    m_pColumnProps = new TYPE_COLUMN_PROPS[ columnCount ];
    if ( m_pColumnProps == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    memset( m_pColumnProps, 0, sizeof ( TYPE_COLUMN_PROPS ) * columnCount );
    m_uNumColumns = columnCount;

    // Allocate the row status fields
    numRecords = m_AuditRecords.GetSize();
    m_pRowStatus = new SQLUSMALLINT[ numRecords ];
    if ( m_pRowStatus == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    memset( m_pRowStatus, 0, sizeof ( SQLUSMALLINT ) * numRecords );

    // Initialise and alloc the field buffers, column-wise
    for ( i = 0; i < m_uNumColumns; i++ )
    {
        length     = 0;
        columnType = RecordSet.GetColumnSqlType( i );
        if ( columnType == SQL_INTEGER )
        {
            length = sizeof ( SQLINTEGER );
            m_pColumnProps[ i ].sqlType    = SQL_INTEGER;
            m_pColumnProps[ i ].TargetType = SQL_C_SLONG;
        }
        else if ( ( columnType == SQL_CHAR        ) ||
                  ( columnType == SQL_VARCHAR     ) ||
                  ( columnType == SQL_LONGVARCHAR )  )
        {
            length = PXSAddSizeT( MAX_COL_SIZE_CHARS, 1 );  // Account for NULL
            length = PXSMultiplySizeT( length, sizeof ( char ) );
            m_pColumnProps[ i ].sqlType    = SQL_VARCHAR;
            m_pColumnProps[ i ].TargetType = SQL_C_CHAR;
        }
        else
        {
            // Wide strings
            length = PXSAddSizeT( MAX_COL_SIZE_CHARS, 1 );  // Account for NULL
            length = PXSMultiplySizeT( length, sizeof ( wchar_t ) );
            m_pColumnProps[ i ].sqlType    = SQL_WVARCHAR;
            m_pColumnProps[ i ].TargetType = SQL_C_WCHAR;
        }

        numBytes = PXSMultiplySizeT( length, numRecords );
        m_pColumnProps[ i ].BufferLength   = PXSCastSizeTToSqlLen( length );
        m_pColumnProps[ i ].TargetValuePtr = new BYTE[ numBytes ];
        if ( m_pColumnProps[ i ].TargetValuePtr == nullptr )
        {
            throw MemoryException( __FUNCTION__ );
        }
        memset( m_pColumnProps[ i ].TargetValuePtr, 0, numBytes );
        m_pColumnProps[ i ].numBytes = numBytes;

        // Allocate StrLen_or_Ind for each column
        m_pColumnProps[ i ].StrLen_or_Ind = new SQLLEN[ numRecords ];
        if ( m_pColumnProps[ i ].StrLen_or_Ind == nullptr )
        {
            throw MemoryException( __FUNCTION__ );
        }
        memset( m_pColumnProps[ i ].StrLen_or_Ind,
                0,
                sizeof ( SQLLEN ) * numRecords );
    }
}

//===============================================================================================//
//  Description:
//      Export the records to the database
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void OdbcExportDialog::ExportRecords()
{
    File   FileObject;
    Odbc   OdbcObject;
    String DbmsName, DatabaseName, ResultMsg, OdbcDriver, ErrorMessage, Text;
    String ConnectionString, ProgressMessage, Password, PortNumber;
    Formatter    Format;
    AuditDatabase Database;

    if ( m_bControlsCreated == false )
    {
        throw FunctionException( L"m_bControlsCreated", __FUNCTION__ );
    }

    if ( m_AuditRecords.GetSize() == 0 )
    {
        throw FunctionException( L"m_AuditRecords", __FUNCTION__ );
    }
    ProgressMessage = PXS_STRING_EMPTY;
    SetProgressMessage( ProgressMessage );

    // Get the database name
    m_DatabaseNameTextField.GetText( &DatabaseName );
    DatabaseName.Trim();
    if ( DatabaseName.IsEmpty() )
    {
        PXSGetResourceString( PXS_IDS_1230_NO_DB_NAME, &ErrorMessage );
        throw SystemException( ERROR_INVALID_DATA, ErrorMessage.c_str(), __FUNCTION__ );
    }

    // If want an access database, ensure have a file path rather than a name
    if ( m_AccessOption.GetState() )
    {
        if ( DatabaseName.IndexOf( PXS_PATH_SEPARATOR, 0 ) == PXS_MINUS_ONE )
        {
            PXSGetResourceString( PXS_IDS_1231_FULL_PATH_REQUIRED, &ErrorMessage );
            throw SystemException( ERROR_BAD_PATHNAME, ErrorMessage.c_str(), __FUNCTION__ );
        }

        // Verify the database exists
        if ( FileObject.Exists( DatabaseName ) == false )
        {
            throw SystemException( ERROR_PATH_NOT_FOUND, DatabaseName.c_str(), __FUNCTION__);
        }

        // Must have an extension as this determines which ODBC driver
        // is used
        if ( ( DatabaseName.EndsWithStringI( L".mdb"   ) == false ) &&
             ( DatabaseName.EndsWithStringI( L".accdb" ) == false )  )
        {
            PXSGetResourceString( PXS_IDS_1239_USE_MDB_OR_ACCDB, &ErrorMessage );
            throw SystemException( ERROR_BAD_PATHNAME, ErrorMessage.c_str(), __FUNCTION__ );
        }
    }
    else
    {
        // Server type database, want a name only
        if ( ( PXS_MINUS_ONE != DatabaseName.IndexOf( '/', 0 ) ) ||
             ( PXS_MINUS_ONE != DatabaseName.IndexOf( ':', 0 ) ) ||
             ( PXS_MINUS_ONE != DatabaseName.IndexOf( PXS_PATH_SEPARATOR, 0) ) )
        {
            PXSGetResourceString( PXS_IDS_1232_USE_DB_NAME_ONLY, &ErrorMessage );
            throw SystemException( ERROR_INVALID_DATA, ErrorMessage.c_str(), __FUNCTION__ );
        }
    }

    // Data looks reasonable so save at class scope
    UpdateConfigurationSettings();

    // Make the connection string
    if ( m_MySQLOption.GetState() )
    {
        m_MySQLComboBox.GetSelectedString( &OdbcDriver );
    }
    else if ( m_PostgreSQLOption.GetState() )
    {
        m_PostgreComboBox.GetSelectedString( &OdbcDriver );
    }
    Password   = PXS_STRING_EMPTY;      // No password
    PortNumber = PXS_STRING_EMPTY;      // No port number
    OdbcObject.MakeConnectionString( m_Settings.DBMS,
                                     OdbcDriver,
                                     m_Settings.ServerName,
                                     m_Settings.DatabaseName,
                                     m_Settings.UID, Password, PortNumber, ConnectionString );
    PXSLogAppInfo1( L"Connection String: '%%1'", ConnectionString );
    Database.Connect( ConnectionString,
                      m_Settings.connectTimeoutSecs, m_Settings.queryTimeoutSecs, m_hWindow );

    // Store user's choices in the configuration
    m_Settings.ServerName = Database.GetServerName();
    m_Settings.UID = Database.GetUserName();

    try
    {
        // Start the progress bar
        PXSGetResourceString( PXS_IDS_1234_EXPORTING_RECORDS, &Text );
        m_ProgressBar.SetText( Text );
        m_ProgressBar.SetPercentage( 0 );
        m_ProgressBar.SetVisible( true );
        m_ProgressBar.RedrawNow();
        m_ProgressBar.SetPercentage( 10 );

        // Do the export
        m_ProgressBar.SetPercentage( 25 );
        WaitCursor Wait;
        ExportRecordsToDatabase( &Database, &ResultMsg );
        m_ProgressBar.SetPercentage( 100 );
        m_ProgressBar.SetVisible( false );
        SetProgressMessage( ResultMsg );
    }
    catch ( const Exception& )
    {
        m_ProgressBar.SetVisible( false );
        throw;
    }
}

//===============================================================================================//
//  Description:
//      Bind the data to the buffers for the Audit_Data table
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void OdbcExportDialog::FillAuditDataBindBuffers( SQLINTEGER auditID, SQLINTEGER computerID )
{
    DWORD   categoryID = 0;
    size_t  i = 0, j = 0, numRecords = 0, numValues = 0, length = 0, offset = 0;
    size_t  dataBytes = 0, maxChars = 0;
    char    szAnsi[ MAX_COL_SIZE_CHARS + 1 ] = { 0 };
    wchar_t szWide[ MAX_COL_SIZE_CHARS + 1 ] = { 0 };
    String  Value;
    Formatter    Format;
    SQLLEN       strLenOrInd = 0;
    SQLPOINTER   pData = nullptr;
    SQLINTEGER   recordOrdinal = 0, sqlInteger = 0;
    AuditRecord  Record;
    StringArray  Values;

    if ( ( auditID == 0 ) || ( computerID == 0 ) )
    {
        throw ParameterException( L"auditID/computerID", __FUNCTION__ );
    }

    if ( ( m_pColumnProps == nullptr ) || ( m_pRowStatus == nullptr ) )
    {
        throw NullException( L"m_pColumnProps/m_pRowStatus", __FUNCTION__ );
    }

    numRecords = m_AuditRecords.GetSize();
    for ( i = 0; i < numRecords; i++ )
    {
        recordOrdinal = PXSAddInt32( recordOrdinal, 1 );
        categoryID = 0;
        Values.RemoveAll();
        Record = m_AuditRecords.Get( i );
        Record.GetCategoryIdAndValues( &categoryID, &Values );
        numValues = Values.GetSize();

        // The first four columns are integers
        pData       = nullptr;
        dataBytes   = 0;
        strLenOrInd = 0;
        for ( j = 0; j < m_uNumColumns; j++ )
        {
            pData = nullptr;
            if ( j == 0 )               // Audit ID
            {
                pData     = &auditID;
                dataBytes = sizeof ( auditID );
            }
            else if ( j == 1 )          // Record Ordinal
            {
                recordOrdinal = PXSAddInt32( recordOrdinal, 1 );
                pData         = &recordOrdinal;
                dataBytes     = sizeof ( recordOrdinal );
            }
            else if ( j == 2 )          // Computer ID
            {
                pData     = &computerID;
                dataBytes = sizeof ( computerID );
            }
            else if ( j == 3 )          // Category ID
            {
                sqlInteger = PXSCastUInt32ToSqlInteger( categoryID );
                pData      = &sqlInteger;
                dataBytes  = sizeof ( sqlInteger );
            }
            else if ( j >= 4 )
            {
                // These are strings, truncate to the column length in bytes.
                strLenOrInd = SQL_NTS;
                if ( ( j - 4 ) < numValues )
                {
                    Value    = Values.Get( j - 4 );
                    maxChars = PXSCastSqlLenToSizeT( m_pColumnProps[ j ].BufferLength );
                    if ( m_pColumnProps[ j ].sqlType == SQL_VARCHAR )
                    {
                        if ( maxChars )
                        {
                            maxChars--;                     // Room for the NULL
                        }
                        Value.Truncate( maxChars );
                        memset( szAnsi, 0, sizeof ( szAnsi ) );
                        Format.StringToAnsi( Value, szAnsi, ARRAYSIZE(szAnsi) );
                        pData     = szAnsi;
                        dataBytes = strlen( szAnsi );
                    }
                    else
                    {
                        if ( maxChars >= sizeof ( wchar_t ) )
                        {
                            maxChars -= sizeof ( wchar_t );  // Room for NULL
                        }
                        Value.Truncate( maxChars / sizeof ( wchar_t ) );
                        memset( szWide, 0, sizeof ( szWide ) );
                        Format.StringToWide( Value, szWide, ARRAYSIZE(szWide) );
                        pData     = szWide;
                        dataBytes = wcslen( szWide );
                        dataBytes = PXSMultiplySizeT( dataBytes, sizeof ( wchar_t ) );
                    }
                }
            }

            if ( pData )
            {
                m_pColumnProps[ j ].StrLen_or_Ind[ i ] = strLenOrInd;
                length = PXSCastSqlLenToSizeT( m_pColumnProps[j].BufferLength );
                offset = PXSMultiplySizeT( i, length );
                if ( PXSAddSizeT(offset, length) > m_pColumnProps[ j ].numBytes)
                {
                    throw SystemException( ERROR_INSUFFICIENT_BUFFER,
                                           PXS_STRING_EMPTY, __FUNCTION__ );
                }
                memcpy( m_pColumnProps[ j ].TargetValuePtr + offset, pData, dataBytes );
            }
            else
            {
                m_pColumnProps[ j ].StrLen_or_Ind[ i ] = SQL_NULL_DATA;
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Free the binding buffers for the Audit_Data table
//
//  Parameters:
//      void
//
//  Returns:
//      void
//===============================================================================================//
void OdbcExportDialog::FreeAuditDataBindBuffers()
{
    size_t i = 0;

    // Columns
    if ( m_pColumnProps )
    {
        // Values
        for ( i = 0; i < m_uNumColumns; i++ )
        {
            if ( m_pColumnProps[ i ].TargetValuePtr )
            {
                delete [] m_pColumnProps[ i ].TargetValuePtr;
            }

            if ( m_pColumnProps[ i ].StrLen_or_Ind )
            {
                delete [] m_pColumnProps[ i ].StrLen_or_Ind;
            }
        }

        // Column array
        delete [] m_pColumnProps;
        m_pColumnProps = nullptr;
    }
    m_uNumColumns = 0;

    // Row Status
    if ( m_pRowStatus )
    {
        delete [] m_pRowStatus;
        m_pRowStatus = nullptr;
    }
}

//===============================================================================================//
//  Description:
//      Set text progress message on the dialog box
//
//  Parameters:
//      ProgressMessage - one-line message
//
//  Returns:
//      void
//===============================================================================================//
void OdbcExportDialog::SetProgressMessage( const String& ProgressMessage )
{
    RECT bounds = { 0, 0, 0, 0 };
    StaticControl Static;

    if ( m_hWindow == nullptr ) return;

    // The static the last one in the collection
    if ( m_idxProgressMessageStatic < m_Statics.GetSize() )
    {
        // Update with new text
        Static = m_Statics.Get( m_idxProgressMessageStatic );
        Static.SetText( ProgressMessage );
        m_Statics.Set( m_idxProgressMessageStatic, Static );
        Static.GetBounds( &bounds );
        InvalidateRect( m_hWindow, &bounds, TRUE );
        RedrawWindow( m_hWindow, &bounds, nullptr, RDW_UPDATENOW );
    }
}

//===============================================================================================//
//  Description:
//      Show the administrator's dialog box
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void OdbcExportDialog::ShowAdminDialog()
{
    String DatabaseName, ErrorMessage, ApplicationName, OdbcDriver;
    DatabaseAdministrationDialog AdminDialog;

    // Database name, make sure have it
    m_DatabaseNameTextField.GetText( &DatabaseName );
    DatabaseName.Trim();
    if ( DatabaseName.IsEmpty() )
    {
        PXSGetResourceString( PXS_IDS_1230_NO_DB_NAME,
                                         &ErrorMessage );
        throw SystemException( ERROR_INVALID_NAME,
                               ErrorMessage.c_str(), __FUNCTION__ );
    }

    // If want to create an access database, ensure have a file path rather
    // than just a name
    if ( m_AccessOption.GetState() )
    {
        if ( DatabaseName.IndexOf( PXS_PATH_SEPARATOR, 0 ) == PXS_MINUS_ONE )
        {
            PXSGetResourceString( PXS_IDS_1231_FULL_PATH_REQUIRED, &ErrorMessage );
            throw SystemException( ERROR_BAD_PATHNAME, ErrorMessage.c_str(), __FUNCTION__ );
        }

        // Must have an extension as this determines which OLD DB driver
        // is used to create the database
        if ( ( DatabaseName.EndsWithStringI( L".mdb"  ) == false ) &&
             ( DatabaseName.EndsWithStringI( L".accdb") == false )  )
        {
            PXSGetResourceString( PXS_IDS_1239_USE_MDB_OR_ACCDB, &ErrorMessage );
            throw SystemException( ERROR_BAD_PATHNAME, ErrorMessage.c_str(), __FUNCTION__ );
        }
    }
    else
    {
        // Server type database, want a name only
        if ( DatabaseName.IndexOf( PXS_PATH_SEPARATOR, 0 ) != PXS_MINUS_ONE )
        {
            PXSGetResourceString( PXS_IDS_1232_USE_DB_NAME_ONLY, &ErrorMessage );
            throw SystemException( ERROR_INVALID_NAME, ErrorMessage.c_str(), __FUNCTION__ );
        }
    }

    // For MySQL and PostgreSQL must have an ODBC driver name
    if ( m_MySQLOption.GetState() )
    {
        m_MySQLComboBox.GetSelectedString( &OdbcDriver );
        if ( OdbcDriver.IsEmpty() )
        {
            throw SystemException( ERROR_BAD_DRIVER, L"MySqlDriver=NULL", __FUNCTION__ );
        }
    }
    else if ( m_PostgreSQLOption.GetState() )
    {
        m_PostgreComboBox.GetSelectedString( &OdbcDriver );
        if ( OdbcDriver.IsEmpty() )
        {
            throw SystemException( ERROR_BAD_DRIVER, L"PostgreSQLDriver=NULL", __FUNCTION__ );
        }
    }

    // Show the dialog
    UpdateConfigurationSettings();
    PXSGetApplicationName( &ApplicationName );
    AdminDialog.SetTitle( ApplicationName );
    AdminDialog.SetConfigurationSettings( m_Settings );
    if ( IDOK == AdminDialog.Create( m_hWindow ) )
    {
        AdminDialog.GetConfigurationSettings( &m_Settings );
    }
}

//===============================================================================================//
//  Description:
//      Update this dialog's settings
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void OdbcExportDialog::UpdateConfigurationSettings()
{
    // DBMS
    if ( m_AccessOption.GetState() )
    {
        m_AccessOption.GetText( &m_Settings.DBMS );
    }
    else if ( m_SqlServerOption.GetState() )
    {
        m_SqlServerOption.GetText( &m_Settings.DBMS );
    }
    else if ( m_MySQLOption.GetState() )
    {
        m_MySQLOption.GetText( &m_Settings.DBMS );
    }
    else if ( m_PostgreSQLOption.GetState() )
    {
        m_PostgreSQLOption.GetText( &m_Settings.DBMS );
    }
    m_MySQLComboBox.GetSelectedString( &m_Settings.MySqlDriver );
    m_PostgreComboBox.GetSelectedString( &m_Settings.PostgreSqlDriver );

    // Database name
    m_DatabaseNameTextField.GetText( &m_Settings.DatabaseName );
    m_Settings.DatabaseName.Trim();
    m_Settings.DatabaseName.Truncate( MAX_PATH );

    // Errors
    m_Settings.maxErrorRate = m_MaxErrorRateSpinner.GetValue();
    PXSLimitUInt32( PXS_DB_MAX_ERROR_RATE_MIN,
                    PXS_DB_MAX_ERROR_RATE_MAX, &m_Settings.maxErrorRate );

    // Affected rows
    m_Settings.maxAffectedRows = m_MaxAffectedRowsSpinner.GetValue();
    PXSLimitUInt32( PXS_DB_MAX_AFFECTED_ROWS_MIN,
                    PXS_DB_MAX_AFFECTED_ROWS_MAX, &m_Settings.maxAffectedRows );

    // Connect timeout
    m_Settings.connectTimeoutSecs = m_ConnectTimeoutSpinner.GetValue();
    PXSLimitUInt32( PXS_DB_CONNECT_TIMEOUT_SECS_MIN,
                    PXS_DB_CONNECT_TIMEOUT_SECS_MAX, &m_Settings.connectTimeoutSecs );

    // Query timeout
    m_Settings.queryTimeoutSecs = m_QueryTimeoutSpinner.GetValue();
    PXSLimitUInt32( PXS_DB_QUERY_TIMEOUT_SECS_MIN,
                    PXS_DB_QUERY_TIMEOUT_SECS_MAX, &m_Settings.queryTimeoutSecs );
}
