///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Database Administration Dialog Class Header Implementation
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
#include "WinAudit/Header Files/DatabaseAdministrationDialog.h"

// 2. C System Files
#include <limits.h>
#include <sqlucode.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/MessageDialog.h"
#include "PxsBase/Header Files/NameValue.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StaticControl.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/SystemInformation.h"
#include "PxsBase/Header Files/WaitCursor.h"

// 5. This Project
#include "WinAudit/Header Files/AccessDatabase.h"
#include "WinAudit/Header Files/AuditData.h"
#include "WinAudit/Header Files/Odbc.h"
#include "WinAudit/Header Files/OdbcRecordSet.h"
#include "WinAudit/Header Files/Resources.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default Constructor
DatabaseAdministrationDialog::DatabaseAdministrationDialog()
                             :m_uIdxStaticProgress( PXS_MINUS_ONE ),
                              m_OdbcDriver(),
                              m_CloseButton(),
                              m_CreateButton(),
                              m_DeleteHistoryButton(),
                              m_RunReportButton(),
                              m_GrantPublicCheckBox(),
                              m_ReportShowSqlCheckBox(),
                              m_UnicodeCheckBox(),
                              m_AccessComboBox(),
                              m_ReportNamesComboBox(),
                              m_ProgressBar(),
                              m_MaxReportRowsSpinner(),
                              m_AuditDatabase(),
                              m_Settings()
{
    try
    {
        SetSize( 440, 370 );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy Constructor - not allowed, no implementation

// Destructor
DatabaseAdministrationDialog::~DatabaseAdministrationDialog()
{
    try
    {
        if ( m_AuditDatabase.IsConnected() )
        {
            if ( m_AuditDatabase.StartedTrans() )
            {
                m_AuditDatabase.RollbackTrans();
            }
            m_AuditDatabase.Disconnect();
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator - not allowed, no implementation

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the settings used by the dialog
//
//  Parameters:
//      Settings - receives the configuration settings
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::GetConfigurationSettings(
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
//      Set the settings used by this dialog
//
//  Parameters:
//      Settings - the configuration settings
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::SetConfigurationSettings(const ConfigurationSettings& Settings )
{
    m_Settings = Settings;

    // Tidy up those settings that pertain to this dialog
    if ( ( m_Settings.connectTimeoutSecs < PXS_DB_CONNECT_TIMEOUT_SECS_MIN ) ||
         ( m_Settings.connectTimeoutSecs > PXS_DB_CONNECT_TIMEOUT_SECS_MAX ) )
    {
        m_Settings.connectTimeoutSecs = PXS_DB_CONNECT_TIMEOUT_SECS_DEF;
    }

    if ( ( m_Settings.queryTimeoutSecs < PXS_DB_QUERY_TIMEOUT_SECS_MIN ) ||
         ( m_Settings.queryTimeoutSecs > PXS_DB_QUERY_TIMEOUT_SECS_MIN ) )
    {
        m_Settings.queryTimeoutSecs = PXS_DB_QUERY_TIMEOUT_SECS_DEF;
    }

    if ( ( m_Settings.reportMaxRecords < PXS_REPORT_MAX_RECORDS_MIN ) ||
         ( m_Settings.reportMaxRecords > PXS_REPORT_MAX_RECORDS_MAX ) )
    {
        m_Settings.queryTimeoutSecs = PXS_REPORT_MAX_RECORDS_DEFAULT;
    }

    m_Settings.DBMS.Trim();
    m_Settings.DatabaseName.Trim();
    m_Settings.LastReportName.Trim();
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
LRESULT DatabaseAdministrationDialog::CommandEvent( WPARAM wParam, LPARAM lParam )
{
    LRESULT result  = 0;                // = Event handled
    WORD    wID     = LOWORD(wParam);   // Item, control, or accelerator id
    String  DatabaseName, HistoryCount;

    // Ensure controls have been created
    if ( m_bControlsCreated == false ) return result;

    // Don't let any exceptions propagate out of event handler
    try
    {
        if ( wID == IDCANCEL )
        {
            EndDialog( m_hWindow, IDCANCEL );
        }
        else if ( IsClickFromButton( m_CloseButton, wParam, lParam ) )
        {
            // Ensure the dialog will close
            try
            {
                UpdateSettings();
            }
            catch ( const Exception& e )
            {
                PXSShowExceptionDialog( e, m_hWindow );
            }
            EndDialog( m_hWindow, IDOK );
        }
        else if ( IsClickFromButton( m_CreateButton, wParam, lParam ) )
        {
            CreateDatabase();
            m_CreateButton.Repaint();
        }
        else if ( IsClickFromButton( m_DeleteHistoryButton, wParam, lParam ) )
        {
            DeleteOldAudits();
            m_DeleteHistoryButton.Repaint();
        }
        else if ( IsClickFromButton( m_RunReportButton, wParam, lParam ) )
        {
            RunReport();
            m_RunReportButton.Repaint();
        }
        else
        {
            result = 1;    // Not handled, return non-zero
        }
    }
    catch ( const Exception& e )
    {
        PXSShowExceptionDialog( e, m_hWindow );
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
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::InitDialogEvent()
{
    int       CONTROL_HEIGHT = 20, LINE_HEIGHT = 23, GAP = 10;
    bool      rtlReading;
    size_t    i = 0, numProviders = 0;
    Font      FontObject;
    SIZE      buttonSize = { 0, 0 };
    SIZE      dialogSize = { 0, 0 };
    RECT      bounds     = { 0, 0, 0, 0 };
    POINT     location   = { 0, 0 };
    DWORD     frameShape = PXS_SHAPE_FRAME;
    String    Text, TempString, AccessProviderID;
    AuditData Audit;
    Formatter Format;
    StaticControl  Static;
    TArray< NameValue > OleDbProviders;

    if ( m_hWindow == nullptr ) return;

    buttonSize.cx = 80;
    buttonSize.cy = 25;
    GetClientSize( &dialogSize );
    rtlReading  = IsRightToLeftReading();
    if ( rtlReading )
    {
        frameShape = PXS_SHAPE_FRAME_RTL;
    }

    // Title Caption, centred
    PXSGetResourceString( PXS_IDS_1110_ADMINISTRATIVE_TASKS, &Text );
    bounds.left   = GAP;
    bounds.right  = dialogSize.cx - GAP;
    bounds.top    = 5;
    bounds.bottom = 30;
    Static.SetBounds( bounds );
    FontObject.SetPointSize( 12, nullptr );
    FontObject.SetBold( true );
    FontObject.SetUnderlined( true );
    FontObject.Create();
    Static.SetFont( FontObject );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    Static.SetText( Text );
    m_Statics.Add( Static );
    Static.Reset();

    bounds.top    = bounds.bottom;
    bounds.bottom = 60;
    Static.SetBounds( bounds );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    Static.SetForeground( PXS_COLOUR_NAVY );
    Text = Format.String2( L"%%1: %%2",
                           m_Settings.DBMS, m_Settings.DatabaseName );
    Static.SetText( Text );
    if ( Text.GetLength() < 30 )
    {
        FontObject.SetPointSize( 10, nullptr );
    }
    else
    {
        FontObject.SetPointSize( 8, nullptr );
    }
    FontObject.SetBold( true );
    FontObject.SetUnderlined( false );
    FontObject.Create();
    Static.SetFont( FontObject );
    m_Statics.Add( Static );
    Static.Reset();

    // Frame for Create Database
    bounds.top    = bounds.bottom;
    bounds.left   = GAP;
    bounds.right  = dialogSize.cx - GAP;
    bounds.bottom = bounds.top + ( 3 * LINE_HEIGHT ) + GAP;
    Static.SetBounds( bounds );
    Static.SetShape( frameShape );
    PXSGetResourceString( PXS_IDS_1100_CREATE_DATABASE, &Text );
    Static.SetText( Text );
    m_Statics.Add( Static );
    Static.Reset();

    // Unicode, Access 2000 and newer is always Unicode so need to present the
    // option. MySql ODBC says "" for SQL_WCHAR and SQL_WVARCHAR so do not allow
    // option. PostgreSQL is ANSI or Unicode when it is created.
    Text = L"Unicode";
    bounds.left   = GAP + GAP;
    bounds.top   += LINE_HEIGHT;
    bounds.right  = dialogSize.cx / 2;
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    m_UnicodeCheckBox.Create( m_hWindow );
    m_UnicodeCheckBox.SetBounds( bounds );
    m_UnicodeCheckBox.SetText( Text );
    m_UnicodeCheckBox.SetState( false );
    m_UnicodeCheckBox.SetVisible( false );
    if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_SQL_SERVER ) == 0 )
    {
        m_UnicodeCheckBox.SetVisible( true );
    }

    // Add check box for WinAudit role, only for SQL Server
    Text = L"GRANT PUBLIC";
    bounds.left   = GAP + GAP;
    bounds.top   += LINE_HEIGHT;
    bounds.right  = dialogSize.cx / 2;
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    m_GrantPublicCheckBox.Create( m_hWindow );
    m_GrantPublicCheckBox.SetBounds( bounds );
    m_GrantPublicCheckBox.SetText( Text );
    m_GrantPublicCheckBox.SetVisible( false );
    if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_SQL_SERVER ) == 0 )
    {
        m_GrantPublicCheckBox.SetVisible( true );
    }

    // Access OLE DB combo list, same location as above as they
    // are mutually exclusive
    m_AccessComboBox.SetBounds( bounds.left,
                                bounds.top, bounds.right - bounds.left, 200 );
    m_AccessComboBox.Create( m_hWindow );
    if ( m_Settings.DatabaseName.EndsWithStringI( L".accdb" ) )
    {
        Audit.GetOleDbProviders( &OleDbProviders );
        numProviders = OleDbProviders.GetSize();
        for ( i = 0; i < numProviders; i++ )
        {
            // Want ACE providers only
            AccessProviderID = OleDbProviders.Get( i ).GetName();
            if ( AccessProviderID.IndexOfI( L"ACE.OLEDB" ) != PXS_MINUS_ONE )
            {
                m_AccessComboBox.Add( AccessProviderID.c_str() );
            }
        }
    }
    else
    {
        // Access 2000
        m_AccessComboBox.Add( L"Microsoft.Jet.OLEDB.4.0" );
    }

    m_AccessComboBox.SetSelectedIndex( 0 );
    m_AccessComboBox.SetVisible( false );
    if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_ACCESS ) == 0 )
    {
        m_AccessComboBox.SetVisible( true );
    }

    // Button to create database
    location.x = dialogSize.cx - ( GAP + GAP + buttonSize.cx );
    location.y = bounds.top - 3;
    m_CreateButton.SetLocation( location );
    m_CreateButton.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1101_CREATE, &Text );
    m_CreateButton.SetText( Text );

    // Frame for data maintenance
    bounds.left   = GAP;
    bounds.right  = dialogSize.cx - GAP;
    bounds.top    = bounds.bottom + LINE_HEIGHT;
    bounds.bottom = bounds.top + ( 2 * LINE_HEIGHT ) + GAP;
    Static.SetBounds( bounds );
    Static.SetShape( frameShape );
    PXSGetResourceString( PXS_IDS_1102_DATA_MAINTENANCE, &Text );
    Static.SetText( Text );
    m_Statics.Add( Static );
    Static.Reset();

    // Message
    bounds.left  += GAP;
    bounds.top   += ( GAP + GAP );
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    bounds.right  = 160;
    Static.SetBounds( bounds );
    PXSGetResourceString( PXS_IDS_1103_DELETE_OLD_AUDITS, &Text );
    Static.SetText( Text );
    m_Statics.Add( Static );
    Static.Reset();

    // Button to delete the old audits
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    location.x = dialogSize.cx - ( GAP + GAP + buttonSize.cx );
    location.y = bounds.top;
    m_DeleteHistoryButton.SetLocation( location );
    m_DeleteHistoryButton.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1104_DELETE, &Text );
    m_DeleteHistoryButton.SetText( Text );

    // Frame for reports
    bounds.left   = GAP;
    bounds.right  = dialogSize.cx - GAP;
    bounds.top    = bounds.bottom + LINE_HEIGHT;
    bounds.bottom = bounds.top + ( 3 * LINE_HEIGHT ) + GAP;
    Static.SetBounds( bounds );
    Static.SetShape( frameShape );
    PXSGetResourceString( PXS_IDS_1105_REPORTS, &Text );
    Static.SetText( Text );
    m_Statics.Add( Static );
    Static.Reset();

    // Label
    bounds.left  += GAP;
    bounds.top   += GAP + GAP;
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    bounds.right  = 125;
    Static.SetBounds( bounds );
    PXSGetResourceString( PXS_IDS_1106_REPORT_NAME, &Text );
    Static.SetText( Text );
    m_Statics.Add( Static );
    Static.Reset();

    // Combo box for report names, will use English for these
    bounds.left   = bounds.right;
    bounds.right  = dialogSize.cx - GAP - GAP;
    m_ReportNamesComboBox.SetBounds( bounds.left,
                                     bounds.top,
                                     bounds.right - bounds.left, 200 );
    m_ReportNamesComboBox.Create( m_hWindow );
    m_ReportNamesComboBox.Add( L"Automatic Updates Status" );
    m_ReportNamesComboBox.Add( L"Counts of Audits by Computer" );
    m_ReportNamesComboBox.Add( L"Counts of Computers by Domain/Workgroup" );
    m_ReportNamesComboBox.Add( L"Counts of Errors by Type" );
    m_ReportNamesComboBox.Add( L"Counts of Installed Software Products" );
    m_ReportNamesComboBox.Add( L"Counts of Memory" );
    m_ReportNamesComboBox.Add( L"Counts of Operating Systems" );
    m_ReportNamesComboBox.Add( L"Counts of Running Services" );
    m_ReportNamesComboBox.Add( L"Hardware Summary" );
    m_ReportNamesComboBox.Add( L"Last Audit Date by Computer" );
    m_ReportNamesComboBox.Add( L"List BIOS Versions" );
    m_ReportNamesComboBox.Add( L"List Displays" );
    m_ReportNamesComboBox.Add( L"List Installed Printers" );
    m_ReportNamesComboBox.Add( L"List Network Adapters" );
    m_ReportNamesComboBox.Add( L"List Physical Disks" );
    m_ReportNamesComboBox.Add( L"List Processors" );
    m_ReportNamesComboBox.Add( L"List Users" );

    // If a report was specified, the select it
    m_ReportNamesComboBox.SetSelectedIndex( 0 );    // Select the first one
    if ( m_ReportNamesComboBox.IndexOf( m_Settings.LastReportName.c_str() ) != PXS_MINUS_ONE )
    {
        m_ReportNamesComboBox.SetSelectedString( m_Settings.LastReportName );
    }

    // Label
    bounds.left   = ( GAP + GAP );
    bounds.top   += ( LINE_HEIGHT + 6 );
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    bounds.right  = 125;
    Static.SetBounds( bounds );
    PXSGetResourceString( PXS_IDS_1107_MAXIMUM_ROWS, &Text );
    Static.SetText( Text );
    m_Statics.Add( Static );
    Static.Reset();

    // Maximum rows
    m_MaxReportRowsSpinner.SetBounds( bounds.right, bounds.top, 65, CONTROL_HEIGHT );
    m_MaxReportRowsSpinner.Create( m_hWindow );
    m_MaxReportRowsSpinner.SetRange( PXS_REPORT_MAX_RECORDS_MIN, PXS_REPORT_MAX_RECORDS_MAX );
    m_MaxReportRowsSpinner.SetValue( m_Settings.reportMaxRecords );

    // Check box to show SQL
    bounds.left  = bounds.right + ( 65 + 25 );
    bounds.right = bounds.left + 100;
    m_ReportShowSqlCheckBox.Create( m_hWindow );
    m_ReportShowSqlCheckBox.SetBounds( bounds );
    PXSGetResourceString( PXS_IDS_1108_SHOW_SQL, &Text );
    m_ReportShowSqlCheckBox.SetText( Text );
    m_ReportShowSqlCheckBox.SetState( m_Settings.reportShowSQL );

    // Button to run report
    location.x = dialogSize.cx - ( GAP + GAP + buttonSize.cx );
    location.y = bounds.top - 3;
    m_RunReportButton.SetLocation( location );
    m_RunReportButton.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_1109_RUN, &Text );
    m_RunReportButton.SetText( Text );

    // Progress bar
    bounds.left   = GAP;
    bounds.top    = dialogSize.cy -
                               GAP - buttonSize.cy - GAP - 5 - CONTROL_HEIGHT;
    bounds.bottom = bounds.top + CONTROL_HEIGHT;
    bounds.right  = dialogSize.cx - GAP;
    m_ProgressBar.Create( m_hWindow );
    m_ProgressBar.SetVisible( false );
    m_ProgressBar.SetBounds( bounds );

    // Invisible label in the same location as the progress bars, store its
    // position in the array
    Static.SetBounds( bounds );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    Static.SetForeground( PXS_COLOUR_NAVY );
    m_Statics.Add( Static );
    m_uIdxStaticProgress = m_Statics.GetSize() - 1;

    // Buttons
    location.y = dialogSize.cy - GAP - buttonSize.cy;

    // Horizontal line
    bounds.left    = GAP;
    bounds.right   = dialogSize.cx - GAP;
    bounds.top     = location.y - GAP;
    bounds.bottom  = bounds.top + 2;
    Static.SetBounds( bounds );
    Static.SetShape( PXS_SHAPE_SUNK );
    m_Statics.Add( Static );
    Static.Reset();

    // Close button
    location.x = dialogSize.cx - ( GAP + GAP + buttonSize.cx );
    m_CloseButton.SetLocation( location );
    m_CloseButton.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_130_CLOSE, &Text );
    m_CloseButton.SetText( Text );

    // Mirror for RTL
    if ( rtlReading )
    {
        RtlStatics();
        m_CloseButton.RtlMirror( dialogSize.cx );
        m_CreateButton.RtlMirror( dialogSize.cx );
        m_DeleteHistoryButton.RtlMirror( dialogSize.cx );
        m_RunReportButton.RtlMirror( dialogSize.cx );
        m_GrantPublicCheckBox.RtlMirror( dialogSize.cx );
        m_ReportShowSqlCheckBox.RtlMirror( dialogSize.cx );
        m_UnicodeCheckBox.RtlMirror( dialogSize.cx );
        m_AccessComboBox.RtlMirror( dialogSize.cx );
        m_ReportNamesComboBox.RtlMirror( dialogSize.cx );
        m_ProgressBar.RtlMirror( dialogSize.cx );
        m_MaxReportRowsSpinner.RtlMirror( dialogSize.cx );
    }

    m_bControlsCreated = true;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Add the SQL statements to create the Access procdures to the
//      Statements array
//
//  Parameters:
//      pStatements - receives the  data definition queries
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::AddCreateAccessProcedures( StringArray* pStatements )
{
    String SqlQuery;

    if ( pStatements == nullptr )
    {
        throw ParameterException( L"pStatements", __FUNCTION__ );
    }

    // For Access the parameter name must be different from the column name, if
    // use '@' like for SQL Server get incorrect behaviour when SELECT
    // Access does not have CURRENT_USER, on 2000+ VARCHAR is Unicode
    // Will post zero for the initial last_audit_id
    SqlQuery = L"CREATE PROCEDURE pxs_sp_insert_computer_master(\r\n"
               L"p_Computer_GUID VARCHAR( 40 ), \r\n"
               L"p_DB_User_Name VARCHAR( 128 ), \r\n"
               L"p_MAC_Address VARCHAR( 24 ),\r\n"
               L"p_Smbios_UUID VARCHAR( 40 ),\r\n"
               L"p_Asset_Tag VARCHAR( 24 ),\r\n"
               L"p_Fully_Qualified_Domain_Name VARCHAR( 252 ),\r\n"
               L"p_Site_Name VARCHAR( 64 ),\r\n"
               L"p_Domain_Name VARCHAR( 64 ),\r\n"
               L"p_Computer_Name VARCHAR( 64 ),\r\n"
               L"p_OS_Product_ID VARCHAR( 32 ),\r\n"
               L"p_Other_Identifier VARCHAR( 32 ),\r\n"
               L"p_WinAudit_GUID VARCHAR( 40 ) )\r\n"
               L"AS\r\n"
               L"INSERT INTO Computer_Master( \r\n"
               L"Computer_GUID, \r\n"
               L"DB_User_Name,\r\n"
               L"Database_Local,\r\n"
               L"Last_Audit_ID,\r\n"
               L"MAC_Address,\r\n"
               L"Smbios_UUID,\r\n"
               L"Asset_Tag,\r\n"
               L"Fully_Qualified_Domain_Name,\r\n"
               L"Site_Name,\r\n"
               L"Domain_Name,\r\n"
               L"Computer_Name,\r\n"
               L"OS_Product_ID,\r\n"
               L"Other_Identifier, \r\n"
               L"WinAudit_GUID \r\n"
               L") VALUES (\r\n"
               L"p_Computer_GUID,\r\n"
               L"p_DB_User_Name,\r\n"
               L"NOW,\r\n"
               L"0,\r\n"
               L"p_MAC_Address,\r\n"
               L"p_Smbios_UUID,\r\n"
               L"p_Asset_Tag,\r\n"
               L"p_Fully_Qualified_Domain_Name,\r\n"
               L"p_Site_Name,\r\n"
               L"p_Domain_Name,\r\n"
               L"p_Computer_Name,\r\n"
               L"p_OS_Product_ID,\r\n"
               L"p_Other_Identifier,\r\n"
               L"p_WinAudit_GUID );\r\n";
    pStatements->Add( SqlQuery );

    // Insert into the Audit_Master table
    SqlQuery = L"CREATE PROCEDURE pxs_sp_insert_audit_master(\r\n"
               L"p_Audit_GUID VARCHAR( 40 ),\r\n"
               L"p_DB_User_Name VARCHAR( 128 ), \r\n"
               L"p_Computer_Local DATETIME,\r\n"
               L"p_Computer_UTC DATETIME,\r\n"
               L"p_Computer_ID INTEGER )\r\n"
               L"AS\r\n"
               L"INSERT INTO Audit_Master( \r\n"
               L"Audit_GUID, \r\n"
               L"DB_User_Name,\r\n"
               L"Database_Local,\r\n"
               L"Database_UTC,\r\n"
               L"Computer_Local,\r\n"
               L"Computer_UTC,\r\n"
               L"Computer_ID\r\n"
               L") VALUES (\r\n"
               L"p_Audit_GUID,\r\n"
               L"p_DB_User_Name,\r\n"
               L"NOW,\r\n"
               L"NULL,\r\n"        // Access does not have a UTC keyword
               L"p_Computer_Local,\r\n"
               L"p_Computer_UTC,\r\n"
               L"p_Computer_ID );\r\n";
    pStatements->Add( SqlQuery );
}

//===============================================================================================//
//  Description:
//      Add the SQL statements to create the Audit_Data table to the
//      Statements array
//
//  Parameters:
//      IntegerKeyword  - the database specific SQL_INTEGER keyword
//      TCharKeyword    - the database specific SQL_CHAR/WCHAR keyword
//      VarTCharKeyword - the database specific SQL_VARCHAR/WVARCHAR keyword
//      pStatements     - receives the  data definition queries
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::AddCreateAuditDataSql( const String& IntegerKeyword,
                                                          const String& TCharKeyword,
                                                          const String& VarTCharKeyword,
                                                          StringArray* pStatements )
{
    size_t    i = 0;
    String    SqlQuery;
    Formatter Format;

    if ( pStatements == nullptr )
    {
        throw ParameterException( L"pStatements", __FUNCTION__ );
    }

    SqlQuery.Allocate( 256 );
    SqlQuery = L"CREATE TABLE Audit_Data( "
               L"Audit_ID %%1 NOT NULL, "
               L"Record_Ordinal %%1 NOT NULL, "
               L"Computer_ID %%1 NULL, "
               L"Category_ID %%1 NULL, ";
    SqlQuery.ReplaceI( L"%%1", IntegerKeyword.c_str() );

    // 25 columns of 255 characters
    for ( i = 0; i < 25; i++ )
    {
        SqlQuery += L"Item_";
        SqlQuery += Format.SizeT( i + 1 );
        SqlQuery += L" ";
        SqlQuery += VarTCharKeyword;
        SqlQuery += L"( 255 ) NULL, ";
    }

    // Next 25 columns of 1 character
    for ( i = 25; i < 50; i++ )
    {
        SqlQuery += L"Item_";
        SqlQuery += Format.SizeT( i + 1 );
        SqlQuery += L" ";
        SqlQuery += TCharKeyword;
        SqlQuery += L"( 1 ) NULL, ";
    }
    SqlQuery += L"CONSTRAINT FK_Audit_Data_Audit_ID FOREIGN KEY "
                L"( Audit_ID ) REFERENCES Audit_Master( Audit_ID ) "
                L"ON DELETE CASCADE, CONSTRAINT PK_Audit_Data PRIMARY KEY"
                L" ( Audit_ID, Record_Ordinal ) ) ";
    pStatements->Add( SqlQuery );

    // Indexes
    pStatements->Add( L"CREATE INDEX idx_AD_Computer_ID "
                      L"ON Audit_Data( Computer_ID )" );
    pStatements->Add( L"CREATE INDEX idx_AD_Category_ID "
                      L"ON Audit_Data( Category_ID )" );

    // Empty view
    pStatements->Add( L"CREATE VIEW v_Audit_Data_Empty AS "
                      L"SELECT * FROM Audit_Data WHERE 1 = 2" );
}

//===============================================================================================//
//  Description:
//      Add the SQL statements to create the Audit_Master table to the
//      Statements array
//
//  Parameters:
//      AutoIncrementKeyword- the database specific auto-increment keyword
//      CharKeyword         - the database specific SQL_CHAR keyword
//      IntegerKeyword      - the database specific SQL_INTEGER keyword
//      TCharKeyword        - the database specific SQL_CHAR/WCHAR keyword
//      VarTCharKeyword     - the database specific SQL_VARCHAR/WVARCHAR keyword
//      pStatements         - receives the  data definition queries
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::AddCreateAuditMasterSql( const String& AutoIncrementKeyword,
                                                            const String& IntegerKeyword,
                                                            const String& TimestampKeyword,
                                                            const String& VarCharKeyword,
                                                            const String& VarTCharKeyword,
                                                            StringArray* pStatements )
{
    String    SqlQuery;
    Formatter Format;

    if ( pStatements == nullptr )
    {
        throw ParameterException( L"pStatements", __FUNCTION__ );
    }

    SqlQuery.Allocate( 256 );
    SqlQuery  = L"CREATE TABLE Audit_Master( ";
    SqlQuery += Format.String1( L"Audit_ID %%1 NOT NULL, "       , AutoIncrementKeyword );
    SqlQuery += Format.String1( L"Audit_GUID %%1( 40 ) NULL, "   , VarCharKeyword );
    SqlQuery += Format.String1( L"DB_User_Name %%1( 128 ) NULL, ", VarTCharKeyword );
    SqlQuery += Format.String1( L"Database_Local %%1 NULL, "     , TimestampKeyword );
    SqlQuery += Format.String1( L"Database_UTC %%1 NULL, "       , TimestampKeyword );
    SqlQuery += Format.String1( L"Computer_Local %%1 NULL, "     , TimestampKeyword );
    SqlQuery += Format.String1( L"Computer_UTC %%1 NULL, "       , TimestampKeyword );
    SqlQuery += Format.String1( L"Computer_ID %%1 NULL, "        , IntegerKeyword );
    SqlQuery += L"CONSTRAINT FK_Audit_Master_Computer_ID "
                L"FOREIGN KEY ( Computer_ID ) REFERENCES "
                L"Computer_Master( Computer_ID ) ON DELETE CASCADE ,";
    SqlQuery += L"CONSTRAINT PK_Audit_Master PRIMARY KEY ( Audit_ID ) ) ";
    pStatements->Add( SqlQuery );

    pStatements->Add( L"CREATE UNIQUE INDEX idx_AM_Audit_GUID "
                      L"ON Audit_Master ( Audit_GUID )" );
    pStatements->Add( L"CREATE INDEX idx_AM_Computer_ID ON "
                      L"Audit_Master ( Computer_ID )" );
}

//===============================================================================================//
//  Description:
//      Add the SQL statements to create the Computer_Master table to the
//      Statements array
//
//  Parameters:
//      AutoIncrementKeyword- the database specific auto-increment keyword
//      IntegerKeyword      - the database specific SQL_INTEGER keyword
//      TimestampKeyword    - the database specific SQL_TYPE_TIMESTAMP keyword
//      WCharKeyword        - the database specific SQL_WCHAR keyword
//      VarTCharKeyword     - the database specific SQL_VARCHAR/WVARCHAR keyword
//      pStatements         - receives the statements
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::AddCreateComputerMasterSql( const String& AutoIncrementKeyword,
                                                               const String& IntegerKeyword,
                                                               const String& TimestampKeyword,
                                                               const String& VarCharKeyword,
                                                               const String& VarTCharKeyword,
                                                               StringArray* pStatements )
{
    String    SqlQuery;
    Formatter Format;

    if ( pStatements == nullptr )
    {
        throw ParameterException( L"pStatements", __FUNCTION__ );
    }

    SqlQuery.Allocate( 256 );
    SqlQuery  = L"CREATE TABLE Computer_Master( ";
    SqlQuery += Format.String1( L"Computer_ID %%1 NOT NULL, "             , AutoIncrementKeyword );
    SqlQuery += Format.String1( L"Computer_GUID %%1( 40 ) NULL, "         , VarCharKeyword );
    SqlQuery += Format.String1( L"DB_User_Name %%1( 128 ) NULL, "         , VarTCharKeyword );
    SqlQuery += Format.String1( L"Last_Audit_ID %%1 NULL, "               , IntegerKeyword );
    SqlQuery += Format.String1( L"Database_Local %%1 NULL, "              , TimestampKeyword );
    SqlQuery += Format.String1( L"MAC_Address %%1( 24 ) NULL, "           , VarCharKeyword );
    SqlQuery += Format.String1( L"Smbios_UUID %%1( 40 ) NULL, "           , VarCharKeyword );
    SqlQuery += Format.String1( L"Asset_Tag %%1( 24 ), "                  , VarCharKeyword );
    SqlQuery += Format.String1( L"Fully_Qualified_Domain_Name %%1(252) NULL, ", VarCharKeyword );
    SqlQuery += Format.String1( L"Site_Name %%1( 64 ) NULL, "             , VarCharKeyword );
    SqlQuery += Format.String1( L"Domain_Name %%1( 64 ) NULL, "           , VarCharKeyword );
    SqlQuery += Format.String1( L"Computer_Name %%1( 64 ) NULL, "         , VarCharKeyword );
    SqlQuery += Format.String1( L"OS_Product_ID %%1( 32 ) NULL, "         , VarCharKeyword );
    SqlQuery += Format.String1( L"Other_Identifier %%1( 32 ) NULL, "      , VarCharKeyword );
    SqlQuery += Format.String1( L"WinAudit_GUID %%1( 40 ) NULL, "         , VarCharKeyword );
    SqlQuery += L"CONSTRAINT PK_Computer_Master PRIMARY KEY(Computer_ID) )";
    pStatements->Add( SqlQuery );

    pStatements->Add( L"CREATE INDEX idx_CM_Last_Audit_ID ON "
                      L"Computer_Master ( Last_Audit_ID )" );
    pStatements->Add( L"CREATE INDEX idx_CM_MAC_Address ON "
                      L"Computer_Master ( MAC_Address )" );
    pStatements->Add( L"CREATE INDEX idx_CM_FQDN ON Computer_Master "
                      L"( Fully_Qualified_Domain_Name )" );
    pStatements->Add( L"CREATE INDEX idx_CM_Computer_Name ON "
                      L"Computer_Master ( Computer_Name )" );
    pStatements->Add( L"CREATE UNIQUE INDEX idx_CM_Computer_GUID ON "
                      L"Computer_Master ( Computer_GUID )" );
}

//===============================================================================================//
//  Description:
//      Add the SQL statements to create the Display_Names table to the
//      Statements array
//
//  Parameters:
//      IntegerKeyword   - the database specific SQL_INTEGER keyword
//      VarTCharKeyword  - the database specific SQL_VARCHAR/WVARCHAR keyword
//      pStatements      - receives the  data definition queries
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::AddCreateDisplayNamesSql( const String& IntegerKeyword,
                                                             const String& VarTCharKeyword,
                                                             StringArray* pStatements )

{
    const  size_t NUM_AUDIT_DATA_COLUMNS = 50;
    size_t    i = 0, j = 0;
    DWORD     itemID = 0, categoryID = 0;
    String    SqlQuery, ColumnNames, DisplayNames, CategoryName;
    Formatter Format;
    AuditData Auditor;

    if ( pStatements == nullptr )
    {
        throw ParameterException( L"pStatements", __FUNCTION__ );
    }

    SqlQuery.Allocate( 256 );
    SqlQuery  = L"CREATE TABLE Display_Names( ";
    SqlQuery += Format.String1( L"Category_ID %%1 NOT NULL, ", IntegerKeyword );
    SqlQuery += Format.String1( L"Category_Name %%1( 128 ) NULL, ", VarTCharKeyword );
    for ( i = 0; i < NUM_AUDIT_DATA_COLUMNS; i++ )
    {
        SqlQuery += L"Item_";
        SqlQuery += Format.SizeT( i + 1 );
        SqlQuery += L" ";
        SqlQuery += VarTCharKeyword;
        SqlQuery += L"( 36 ) NULL, ";
    }
    SqlQuery += L"CONSTRAINT PK_Display_Names PRIMARY KEY(Category_ID) ) ";
    pStatements->Add( SqlQuery );

    // Insert Items into the Display_Names table
    ColumnNames.Allocate( 4096 );
    DisplayNames.Allocate( 4096 );
    for ( i = 0; i < ARRAYSIZE( PXS_DATA_CATEGORY_PROPERTIES ); i++ )
    {
        ColumnNames  = PXS_STRING_EMPTY;
        DisplayNames = PXS_STRING_EMPTY;
        CategoryName = PXS_STRING_EMPTY;
        categoryID   = PXS_DATA_CATEGORY_PROPERTIES[ i ].categoryID;
        Auditor.GetCategoryName( categoryID, &CategoryName );
        for ( j = 0; j < ARRAYSIZE( PXS_AUDIT_ITEMS ); j++ )
        {
            itemID = PXS_AUDIT_ITEMS[ j ].itemID;
            if ( ( itemID > categoryID ) &&
                 ( itemID < ( categoryID + PXS_CATEGORY_INTERVAL ) ) )
            {
                if ( ColumnNames.GetLength() )
                {
                    ColumnNames += L", ";
                }
                ColumnNames += L"Item_";
                ColumnNames += Format.UInt32( itemID % PXS_CATEGORY_INTERVAL );

                if ( DisplayNames.GetLength() )
                {
                    DisplayNames += L", ";
                }
                DisplayNames += L"'";
                DisplayNames += PXS_AUDIT_ITEMS[ j ].pszName;
                DisplayNames += L"'";
            }
        }

        // SMBIOS columns
        for ( j = 0; j < ARRAYSIZE( SMBIOS_SPECIFICATION ); j++ )
        {
            itemID = SMBIOS_SPECIFICATION[ j ].itemID;
            if ( ( itemID > categoryID ) &&
                 ( itemID < ( categoryID + PXS_CATEGORY_INTERVAL ) ) )
            {
                if ( ColumnNames.GetLength() )
                {
                    ColumnNames += L", ";
                }
                ColumnNames += L"Item_";
                ColumnNames += Format.UInt32( itemID % PXS_CATEGORY_INTERVAL );

                if ( DisplayNames.GetLength() )
                {
                    DisplayNames += L", ";
                }
                DisplayNames += L"'";
                DisplayNames += SMBIOS_SPECIFICATION[ j ].szName;
                DisplayNames += L"'";
            }
        }

        // Make the SQL
        if ( ColumnNames.GetLength() )
        {
            SqlQuery  = L"INSERT INTO Display_Names";
            SqlQuery += L" ( Category_ID, Category_Name, ";
            SqlQuery += ColumnNames;
            SqlQuery += L" ) VALUES ( ";
            SqlQuery += Format.UInt32( categoryID );
            SqlQuery += L", '";
            SqlQuery += CategoryName;
            SqlQuery += L"', ";
            SqlQuery += DisplayNames;
            SqlQuery += L" )";
            pStatements->Add( SqlQuery );
        }
    }
}

//===============================================================================================//
//  Description:
//      Add the SQL statements to create the PostgreSQL and SQL Server GRANTs
//      to the Statements array
//
//  Parameters:
//      pStatements - receives the  data definition queries
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::AddCreateGrants( StringArray* pStatements )
{
    String    SqlQuery, UserName, Schema;
    Formatter Format;

    if ( pStatements == nullptr )
    {
        throw ParameterException( L"pStatements", __FUNCTION__ );
    }

    // CREATE ROLE is new for SQL Server 2005.
    UserName = m_AuditDatabase.GetUserName();
    if ( UserName.GetLength() )
    {
        Schema = Format.String1( L"[%%1].", UserName );
    }
    SqlQuery = Format.String1( L"GRANT INSERT ON %%1[Computer_Master] TO PUBLIC", Schema );
    pStatements->Add( SqlQuery );

    SqlQuery = Format.String1( L"GRANT SELECT ON %%1[Computer_Master] TO PUBLIC", Schema );
    pStatements->Add( SqlQuery );

    SqlQuery = Format.String1( L"GRANT UPDATE( Last_Audit_ID ) ON %%1[Computer_Master] TO PUBLIC",
                               Schema );
    pStatements->Add( SqlQuery );

    SqlQuery = Format.String1( L"GRANT INSERT ON %%1[Audit_Master] TO PUBLIC", Schema );
    pStatements->Add( SqlQuery );

    SqlQuery = Format.String1( L"GRANT SELECT ON %%1[Audit_Master]("
                               L"Audit_ID, Audit_GUID) TO PUBLIC", Schema);
    pStatements->Add( SqlQuery );

    SqlQuery = Format.String1( L"GRANT INSERT ON %%1[Audit_Data] TO PUBLIC", Schema );
    pStatements->Add( SqlQuery );

    SqlQuery = Format.String1( L"GRANT INSERT ON %%1[v_Audit_Data_Empty] TO PUBLIC", Schema );
    pStatements->Add( SqlQuery );

    SqlQuery = Format.String1( L"GRANT SELECT ON %%1[v_Audit_Data_Empty] TO PUBLIC", Schema );
    pStatements->Add( SqlQuery );

    SqlQuery = Format.String1( L"GRANT EXECUTE ON %%1[pxs_sp_insert_computer_master] TO PUBLIC",
                               Schema );
    pStatements->Add( SqlQuery );

    SqlQuery = Format.String1( L"GRANT EXECUTE ON %%1[pxs_sp_insert_audit_master] TO PUBLIC",
                               Schema );
    pStatements->Add( SqlQuery );
}

//===============================================================================================//
//  Description:
//      Add the SQL statements to create the MySQL procedures to the
//      Statements array
//
//  Parameters:
//      pStatements - receives the  data definition queries
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::AddCreateMySqlProcedures(
                                                      StringArray* pStatements )
{
    String  SqlQuery;

    if ( pStatements == nullptr )
    {
        throw ParameterException( L"pStatements", __FUNCTION__ );
    }

    // MySQL has UTC_TIMSTAMP from v4.1.1
    // Will post zero for the initial last_audit_id
    SqlQuery.Allocate( 256 );
    SqlQuery = L"CREATE PROCEDURE pxs_sp_insert_computer_master(\r\n"
               L"p_Computer_GUID VARCHAR( 40 ),\r\n"
               L"p_DB_User_Name VARCHAR( 128 ), \r\n"
               L"p_MAC_Address VARCHAR( 24 ),\r\n"
               L"p_Smbios_UUID VARCHAR( 40 ),\r\n"
               L"p_Asset_Tag VARCHAR( 24 ),\r\n"
               L"p_Fully_Qualified_Domain_Name VARCHAR( 252 ),\r\n"
               L"p_Site_Name VARCHAR( 64 ),\r\n"
               L"p_Domain_Name VARCHAR( 64 ),\r\n"
               L"p_Computer_Name VARCHAR( 64 ),\r\n"
               L"p_OS_Product_ID VARCHAR( 32 ),\r\n"
               L"p_Other_Identifier VARCHAR( 32 ),\r\n"
               L"p_WinAudit_GUID VARCHAR( 40 ) )\r\n"
               L"BEGIN\r\n"
               L"INSERT INTO Computer_Master( \r\n"
               L"Computer_GUID, \r\n"
               L"DB_User_Name,\r\n"
               L"Database_Local,\r\n"
               L"Last_Audit_ID,\r\n"
               L"MAC_Address,\r\n"
               L"Smbios_UUID,\r\n"
               L"Asset_Tag,\r\n"
               L"Fully_Qualified_Domain_Name,\r\n"
               L"Site_Name,\r\n"
               L"Domain_Name,\r\n"
               L"Computer_Name,\r\n"
               L"OS_Product_ID,\r\n"
               L"Other_Identifier, \r\n"
               L"WinAudit_GUID \r\n"
               L") VALUES (\r\n"
               L"p_Computer_GUID,\r\n"
               L"CURRENT_USER,\r\n"
               L"CURRENT_TIMESTAMP,\r\n"
               L"0,\r\n"               // Initial value for last_audit_id
               L"p_MAC_Address,\r\n"
               L"p_Smbios_UUID,\r\n"
               L"p_Asset_Tag,\r\n"
               L"p_Fully_Qualified_Domain_Name,\r\n"
               L"p_Site_Name,\r\n"
               L"p_Domain_Name,\r\n"
               L"p_Computer_Name,\r\n"
               L"p_OS_Product_ID,\r\n"
               L"p_Other_Identifier,\r\n"
               L"p_WinAudit_GUID );\r\n"
               L"END\r\n";
    pStatements->Add( SqlQuery );

    SqlQuery = L"CREATE PROCEDURE pxs_sp_insert_audit_master(\r\n"
               L"p_Audit_GUID VARCHAR( 40 ),\r\n"
               L"p_DB_User_Name VARCHAR( 128 ), \r\n"
               L"p_Computer_Local DATETIME,\r\n"
               L"p_Computer_UTC DATETIME,\r\n"
               L"p_Computer_ID INTEGER )\r\n"
               L"BEGIN\r\n"
               L"INSERT INTO Audit_Master( \r\n"
               L"Audit_GUID, \r\n"
               L"DB_User_Name,\r\n"
               L"Database_Local,\r\n"
               L"Database_UTC,\r\n"
               L"Computer_Local,\r\n"
               L"Computer_UTC,\r\n"
               L"Computer_ID \r\n"
               L") VALUES (\r\n"
               L"p_Audit_GUID,\r\n"
               L"CURRENT_USER,\r\n"    // Override procedure's input
               L"CURRENT_TIMESTAMP,\r\n"
               L"UTC_TIMESTAMP(),\r\n"
               L"p_Computer_Local,\r\n"
               L"p_Computer_UTC,\r\n"
               L"p_Computer_ID );\r\n"
               L"END\r\n";
    pStatements->Add( SqlQuery );
}

//===============================================================================================//
//  Description:
//      Add the SQL statements to create the MySQL procedures to the
//      Statements array
//
//  Parameters:
//      pStatements - receives the  data definition queries
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::AddCreateSqlServerProcedures( StringArray* pStatements )
{
    String  SqlQuery;

    if ( pStatements == nullptr )
    {
        throw ParameterException( L"pStatements", __FUNCTION__ );
    }

    // sp_executesql requires on v7
    SqlQuery.Allocate( 256 );
    SqlQuery = L"CREATE PROCEDURE [dbo].[pxs_sp_insert_computer_master]"
               L"(\r\n"
               L"@Computer_GUID VARCHAR( 40 ), \r\n"
               L"@DB_User_Name NVARCHAR( 128 ),\r\n"
               L"@MAC_Address VARCHAR( 24 ),\r\n"
               L"@Smbios_UUID VARCHAR( 40 ),\r\n"
               L"@Asset_Tag VARCHAR( 24 ),\r\n"
               L"@Fully_Qualified_Domain_Name VARCHAR( 252 ),\r\n"
               L"@Site_Name VARCHAR( 64 ),\r\n"
               L"@Domain_Name VARCHAR( 64 ),\r\n"
               L"@Computer_Name VARCHAR( 64 ),\r\n"
               L"@OS_Product_ID VARCHAR( 32 ),\r\n"
               L"@Other_Identifier VARCHAR( 32 ),\r\n"
               L"@WinAudit_GUID VARCHAR( 40 ) )\r\n"
               L"AS\r\n"
               L"\r\n"
               L"DECLARE @SQL NVARCHAR( 4000 );\r\n"
               L"\r\n"
               L"\r\n"
               L"SET @SQL = N'INSERT INTO [dbo].[Computer_Master]( \r\n"
               L"Computer_GUID, \r\n"
               L"DB_User_Name,\r\n"
               L"Database_Local,\r\n"
               L"Last_Audit_ID,\r\n"
               L"MAC_Address,\r\n"
               L"Smbios_UUID,\r\n"
               L"Asset_Tag,\r\n"
               L"Fully_Qualified_Domain_Name,\r\n"
               L"Site_Name,\r\n"
               L"Domain_Name,\r\n"
               L"Computer_Name,\r\n"
               L"OS_Product_ID,\r\n"
               L"Other_Identifier, \r\n"
               L"WinAudit_GUID \r\n"
               L") VALUES (\r\n"
               L"@Computer_GUID, \r\n"
               L"USER_NAME(),\r\n"     // Override input to procedure
               L"GETDATE(),\r\n"       // Override input to procedure
               L"0,\r\n"               // Initial value for last_audit_id
               L"@MAC_Address,\r\n"
               L"@Smbios_UUID,\r\n"
               L"@Asset_Tag,\r\n"
               L"@Fully_Qualified_Domain_Name,\r\n"
               L"@Site_Name,\r\n"
               L"@Domain_Name,\r\n"
               L"@Computer_Name,\r\n"
               L"@OS_Product_ID,\r\n"
               L"@Other_Identifier,\r\n"
               L"@WinAudit_GUID )'\r\n"
               L"\r\n"
               L"EXECUTE sp_executesql @SQL,\r\n"
               L"N'@Computer_GUID VARCHAR( 40 ),"
               L"  @MAC_Address VARCHAR( 24 ), "
               L"  @Smbios_UUID VARCHAR( 40 ), "
               L"  @Asset_Tag VARCHAR( 24 ), "
               L"  @Fully_Qualified_Domain_Name VARCHAR( 252 ), "
               L"  @Site_Name VARCHAR( 64 ), "
               L"  @Domain_Name VARCHAR( 64 ), "
               L"  @Computer_Name VARCHAR( 64 ), "
               L"  @OS_Product_ID VARCHAR( 32 ), "
               L"  @Other_Identifier VARCHAR( 32 ) ,"
               L"  @WinAudit_GUID VARCHAR( 40 )',\r\n"
               L"  @Computer_GUID, \r\n"
               L"  @MAC_Address, \r\n"
               L"  @Smbios_UUID, \r\n"
               L"  @Asset_Tag, \r\n"
               L"  @Fully_Qualified_Domain_Name, \r\n"
               L"  @Site_Name, \r\n"
               L"  @Domain_Name, \r\n"
               L"  @Computer_Name, \r\n"
               L"  @OS_Product_ID, \r\n"
               L"  @Other_Identifier, \r\n"
               L"  @WinAudit_GUID \r\n"
               L"\r\n";
    pStatements->Add( SqlQuery );

    SqlQuery = L"CREATE PROCEDURE [dbo].[pxs_sp_insert_audit_master]"
               L"(\r\n"
               L"@Audit_GUID VARCHAR( 40 ), \r\n"
               L"@DB_User_Name NVARCHAR( 128 ),\r\n"
               L"@Computer_Local DATETIME,\r\n"
               L"@Computer_UTC DATETIME,\r\n"
               L"@Computer_ID INT )\r\n"
               L"AS\r\n"
               L"\r\n"
               L"DECLARE @SQL NVARCHAR( 4000 );\r\n"
               L"\r\n"
               L"\r\n"
               L"SET @SQL = N'INSERT INTO [dbo].[Audit_Master]( \r\n"
               L"Audit_GUID, \r\n"
               L"DB_User_Name,\r\n"
               L"Database_Local,\r\n"
               L"Database_UTC,\r\n"
               L"Computer_Local,\r\n"
               L"Computer_UTC,\r\n"
               L"Computer_ID \r\n"
               L") VALUES (\r\n"
               L"@Audit_GUID,\r\n"
               L"USER_NAME(),\r\n"      // Override input to procedure
               L"GETDATE(),\r\n"
               L"GETUTCDATE(),\r\n"
               L"@Computer_Local,\r\n"
               L"@Computer_UTC,\r\n"
               L"@Computer_ID )'\r\n"
               L"\r\n"
               L"EXECUTE sp_executesql @SQL,\r\n"
               L"N'@Audit_GUID VARCHAR( 40 ), "
               L"  @Computer_Local DATETIME, "
               L"  @Computer_UTC DATETIME, "
               L"  @Computer_ID INT',\r\n"
               L"  @Audit_GUID, \r\n"
               L"  @Computer_Local, \r\n"
               L"  @Computer_UTC, \r\n"
               L"  @Computer_ID\r\n"
               L"\r\n";
    pStatements->Add( SqlQuery );
}

//===============================================================================================//
//  Description:
//      Add the SQL statements to create the views to the Statements array
//
//  Parameters:
//      SchemaName - the schema name to which the views belong. Include the
//                   square brackets, e.g. [dbo]
//      pStatements- receives the  data definition queries
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::AddCreateViews( const String& SchemaName,
                                                   StringArray* pStatements )

{
    size_t i = 0, j = 0;
    DWORD  categoryID = 0, itemID = 0;
    String SqlQuery, ColumnNames, ViewName, CategoryName, DisplayName;
    Formatter Format;
    AuditData Auditor;

    if ( pStatements == nullptr )
    {
        throw ParameterException( L"pStatements", __FUNCTION__ );
    }

    for ( i = 0; i < ARRAYSIZE( PXS_DATA_CATEGORY_PROPERTIES ); i++ )
    {
        ColumnNames = PXS_STRING_EMPTY;
        categoryID  = PXS_DATA_CATEGORY_PROPERTIES[ i ].categoryID;
        if ( categoryID >= PXS_CATEGORY_SYSTEM_OVERVIEW )   // Skip groupings
        {
            CategoryName = PXS_STRING_EMPTY;
            Auditor.GetCategoryName( categoryID, &CategoryName );
            ViewName  = L"v_";
            ViewName += CategoryName;
            ViewName.ReplaceChar( ' ', '_' );
            ViewName.ReplaceChar( '/', '_' );
            ViewName.ReplaceChar( '-', '_' );

            // Oops! Two categories names are duplicated, concatenate the
            // categoryID to differentiate them
            if ( ViewName.CompareI( L"v_Processor" ) == 0 )
            {
                ViewName += L"_";
                ViewName += Format.UInt32( categoryID );
            }

            // Always start with these
            ColumnNames  = L"\tAudit_ID as AuditID, ";
            ColumnNames += L"\tRecord_Ordinal as RecordID, ";
            ColumnNames += L"\tcomputer_ID as ComputerID";

            for ( j = 0; j < ARRAYSIZE( PXS_AUDIT_ITEMS ); j++ )
            {
                itemID = PXS_AUDIT_ITEMS[ j ].itemID;
                if ( ( itemID > categoryID ) &&
                     ( itemID < ( categoryID + PXS_CATEGORY_INTERVAL ) ) )
                {
                    DisplayName = PXS_AUDIT_ITEMS[ j ].pszName;
                    DisplayName.ReplaceChar( ' ', '_' );
                    DisplayName.ReplaceChar( '-', '_' );
                    DisplayName.ReplaceChar( '/', '_' );
                    DisplayName.ReplaceChar( '.', PXS_STRING_EMPTY );
                    DisplayName.ReplaceChar( '!', PXS_STRING_EMPTY );
                    DisplayName.ReplaceChar( '+', PXS_STRING_EMPTY );
                    DisplayName.Truncate( 32 );
                    DisplayName.Trim();

                    ColumnNames += L", ";
                    ColumnNames += L"Item_";
                    ColumnNames +=Format.UInt32(itemID % PXS_CATEGORY_INTERVAL);
                    ColumnNames += L" AS ";
                    ColumnNames += DisplayName;
                }
            }

            for ( j = 0; j < ARRAYSIZE( SMBIOS_SPECIFICATION ); j++ )
            {
                itemID = SMBIOS_SPECIFICATION[ j ].itemID;
                if ( ( itemID > categoryID ) &&
                     ( itemID < ( categoryID + PXS_CATEGORY_INTERVAL ) ) )
                {
                    DisplayName = SMBIOS_SPECIFICATION[ j ].szName;
                    DisplayName.ReplaceChar( ' ', '_' );
                    DisplayName.ReplaceChar( '-', '_' );
                    DisplayName.ReplaceChar( '/', '_' );
                    DisplayName.ReplaceChar( '.', '_' );    // e.g. 1.44MB
                    DisplayName.ReplaceChar( '!', PXS_STRING_EMPTY );
                    DisplayName.ReplaceChar( '+', PXS_STRING_EMPTY );
                    DisplayName.Truncate( 32 );
                    DisplayName.Trim();

                    ColumnNames += L", ";
                    ColumnNames += L"Item_";
                    ColumnNames +=Format.UInt32(itemID % PXS_CATEGORY_INTERVAL);
                    ColumnNames += L" AS ";
                    ColumnNames += DisplayName;
                }
            }
            SqlQuery  = L"CREATE VIEW ";
            SqlQuery += SchemaName;
            SqlQuery += ViewName;
            SqlQuery += L" AS \r\n";
            SqlQuery += L"SELECT \r\n";
            SqlQuery += ColumnNames;
            SqlQuery += PXS_STRING_CRLF;
            SqlQuery += L"FROM Audit_Data AS \r\n";
            SqlQuery += ViewName;
            SqlQuery += PXS_STRING_CRLF;
            SqlQuery += L"WHERE Category_ID = ";
            SqlQuery += Format.UInt32( categoryID );
            pStatements->Add( SqlQuery );
        }
    }
}


//===============================================================================================//
//  Description:
//      Connect to the database, if not already connected
//
//  Parameters:
//      serverOnly - connect to the server only, i.e. not to a named database
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::ConnectDB( bool serverOnly )
{
    Odbc   OdbcObject;
    String ConnectionString, OdbcDriver, DatabaseName, Password, PortNumber;

    if ( m_AuditDatabase.IsConnected() )
    {
        return;     // Nothing to do
    }

    if ( m_Settings.DatabaseName.IsEmpty() )
    {
        throw FunctionException( L"m_Settings.DatabaseName", __FUNCTION__ );
    }

    if ( m_Settings.DBMS.IsEmpty() )
    {
        throw FunctionException( L"m_Settings.DBMS", __FUNCTION__ );
    }

    // For MysQL and PostgreSQL an ODBC driver needs to be specified
    if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_MYSQL ) == 0 )
    {
        if ( m_Settings.MySqlDriver.IsEmpty() )
        {
            throw SystemException( ERROR_BAD_DRIVER, L"MySqlDriver=NULL", __FUNCTION__ );
        }
        OdbcDriver = m_Settings.MySqlDriver;
    }
    else if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_POSTGRE_SQL ) == 0 )
    {
        if ( m_Settings.PostgreSqlDriver.IsEmpty() )
        {
            throw SystemException( ERROR_BAD_DRIVER, L"PostgreSqlDriver=NULL", __FUNCTION__ );
        }
        OdbcDriver = m_Settings.PostgreSqlDriver;
    }

    // Make a connection string, use "" if connecting to the server but
    // not using a specific database.
    if ( serverOnly )
    {
        // For PostgreSQL need the database name otherwise
        // the driver's GUI prompt returns SQL_NO_DATA
        if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_POSTGRE_SQL ) == 0 )
        {
            DatabaseName = m_Settings.DatabaseName;
        }
        else
        {
            DatabaseName = PXS_STRING_EMPTY;
        }
    }
    else
    {
        DatabaseName = m_Settings.DatabaseName;
    }
    Password   = PXS_STRING_EMPTY;      // No password
    PortNumber = PXS_STRING_EMPTY;      // No port number
    OdbcObject.MakeConnectionString( m_Settings.DBMS,
                                     OdbcDriver,
                                     m_Settings.ServerName,
                                     DatabaseName,
                                     m_Settings.UID , Password, PortNumber, ConnectionString );
    PXSLogAppInfo1( L"In connection string: '%%1'", ConnectionString );
    m_AuditDatabase.Connect( ConnectionString,
                             m_Settings.connectTimeoutSecs,
                             m_Settings.queryTimeoutSecs, m_hWindow );

    // Store the user's choices at class scope
    m_Settings.ServerName = m_AuditDatabase.GetServerName();
    m_Settings.UID        = m_AuditDatabase.GetUserName();
}

//===============================================================================================//
//  Description:
//      Create the Access database. Requires Jet4/Access 2000 or newer.
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::CreateAccessDatabase()
{
    File   FileObject;
    DWORD  percentage = 0;
    size_t i = 0, numStatements = 0;
    String AutoIncrementKeyword, CharKeyword, SmallIntKeyword, IntegerKeyword;
    String TimestampKeyword, VarCharKeyword, WCharKeyword, WVarCharKeyword;
    String TCharKeyword, VarTCharKeyword, ProviderID, SqlQuery;
    String SchemaName, UserName, Password, ProgressMessage;
    StringArray Statements;
    AccessDatabase Access;

    // Must have an extension as this determines the OLEDB driver
    if ( ( m_Settings.DatabaseName.EndsWithStringI( L".mdb"   ) == false ) &&
         ( m_Settings.DatabaseName.EndsWithStringI( L".accdb" ) == false )  )
    {
        throw SystemException( ERROR_BAD_PATHNAME, m_Settings.DatabaseName.c_str(), __FUNCTION__ );
    }

    // Prevent overwrite
    if ( FileObject.Exists( m_Settings.DatabaseName ) )
    {
        throw SystemException( ERROR_FILE_EXISTS, m_Settings.DatabaseName.c_str(), __FUNCTION__ );
    }
    GetDatabaseSpecificKeywords( &AutoIncrementKeyword,
                                 &CharKeyword,
                                 &IntegerKeyword,
                                 &SmallIntKeyword,
                                 &TimestampKeyword,
                                 &VarCharKeyword, &WCharKeyword, &WVarCharKeyword );
    // Access is Unicode
    TCharKeyword    = WCharKeyword;
    VarTCharKeyword = WVarCharKeyword;

    // Make the statements to create the database
    AddCreateComputerMasterSql( AutoIncrementKeyword,
                                IntegerKeyword,
                                TimestampKeyword, VarCharKeyword, VarTCharKeyword, &Statements );
    AddCreateAuditMasterSql( AutoIncrementKeyword,
                             IntegerKeyword,
                             TimestampKeyword, VarCharKeyword, VarTCharKeyword, &Statements );
    AddCreateAuditDataSql( IntegerKeyword, TCharKeyword, VarTCharKeyword, &Statements );
    AddCreateDisplayNamesSql( IntegerKeyword, VarTCharKeyword, &Statements );
    AddCreateAccessProcedures( &Statements );
    AddCreateViews( SchemaName, &Statements );

    // Create the database's objects
    WaitCursor Cursor;
    m_AccessComboBox.GetSelectedString( &ProviderID );
    Access.OleDbSetProvider( ProviderID );
    Access.OleDbCreate( m_Settings.DatabaseName );
    try
    {
        m_ProgressBar.SetPercentage( 0 );
        m_ProgressBar.SetVisible( true );
        m_ProgressBar.RedrawNow();

        UserName = L"admin";
        Password = PXS_STRING_EMPTY;
        Access.OleDbConnect( m_Settings.DatabaseName, UserName, Password );
        numStatements = Statements.GetSize();
        for ( i = 0; i < numStatements; i++ )
        {
            SqlQuery = Statements.Get( i );
            Access.OleDbExecute( SqlQuery );
            percentage = PXSCastSizeTToUInt32( ( i * 100 ) / numStatements );
            m_ProgressBar.SetPercentage( percentage );
        }
        Access.OleDbDisconnect();
        m_ProgressBar.SetPercentage( 100 );
    }
    catch ( const Exception& )
    {
        m_ProgressBar.SetVisible( false );
        throw;
    }
    m_ProgressBar.SetVisible( false );
    PXSGetResourceString( PXS_IDS_1238_DATABASE_WAS_CREATED, &ProgressMessage );
    SetProgressMessage( ProgressMessage );
}

//===============================================================================================//
//  Description:
//      Create the database
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::CreateDatabase()
{
    String    ResourceString, ApplicationName, Message;
    Formatter Format;

    PXSLogAppInfo( L"Creating the WinAudit database." );
    if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_ACCESS ) == 0 )
    {
        CreateAccessDatabase();
    }
    else
    {
        // Server database. For PostgreSQL, the ODBC connect dialog needs all
        // fields otherwise it returns SQL_NO_DATA. This means an empty
        // database must have already been created.
        if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_POSTGRE_SQL ) == 0 )
        {
            PXSGetResourceString1( PXS_IDS_1243_CONFIRM_EMPTY_DB,
                                   m_Settings.DatabaseName, &ResourceString );
            PXSGetApplicationName( &ApplicationName );
            if ( IDNO == MessageBox( m_hWindow,
                                     ResourceString.c_str(), ApplicationName.c_str(), MB_YESNO ) )
            {
                return;
            }
        }
        CreateServerDatabase();
    }
    PXSLogAppInfo1( L"Created database '%%1'.", m_Settings.DatabaseName );
}

//===============================================================================================//
//  Description:
//      Create the WindAudit MySql, PostgreSQL or SQL Server Database.
//
//  Parameters:
//      None
//
//  Remarks:
//      For PostgreSQL, the empty database must already exist
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::CreateServerDatabase()
{
    DWORD  percentage = 0;
    size_t i = 0, numStatements = 0;
    String AutoIncrementKeyword, CharKeyword, SmallIntKeyword, IntegerKeyword;
    String TimestampKeyword, VarCharKeyword, WCharKeyword, WVarCharKeyword;
    String TCharKeyword, VarTCharKeyword, SqlQuery, UserName, SchemaName;
    String ProgressMessage;
    Formatter   Format;
    StringArray Statements;

    ConnectDB( true );  // Server connect only
    GetDatabaseSpecificKeywords( &AutoIncrementKeyword,
                                 &CharKeyword,
                                 &IntegerKeyword,
                                 &SmallIntKeyword,
                                 &TimestampKeyword,
                                 &VarCharKeyword, &WCharKeyword, &WVarCharKeyword );
    // ANSI/Unicode. Access 2000 and nwer is always Unicode
    // MySql ODBC says "" for SQL_WCHAR and SQL_WVARCHAR
    // PostgreSQL is either ANSI/Unicode on DB creation. It uses char and
    // varchar for SQL_WCHAR and SQL_WVARCHAR respectively even when database
    // is created with a code page of UTF-8 and using the Unicode ODBC driver
    if ( m_UnicodeCheckBox.GetState() )
    {
        TCharKeyword    = WCharKeyword;
        VarTCharKeyword = WVarCharKeyword;
    }
    else
    {
        TCharKeyword    = CharKeyword;
        VarTCharKeyword = VarCharKeyword;
    }

    // Create the database. Will do this in auto-commit mode as usually cannot
    // use CREATE DATABASE in a transaction
    if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_MYSQL ) == 0 )
    {
        SqlQuery = L"CREATE DATABASE ";
        SqlQuery += m_Settings.DatabaseName;
        m_AuditDatabase.ExecuteDirect( SqlQuery );

        SqlQuery = L"USE ";
        SqlQuery += m_Settings.DatabaseName;
        m_AuditDatabase.ExecuteDirect( SqlQuery );
    }
    else if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_SQL_SERVER ) == 0 )
    {
        // CREATE ROLE is new for SQL Server 2005.
        UserName = m_AuditDatabase.GetUserName();
        if ( UserName.GetLength() )
        {
            SchemaName = Format.String1( L"[%%1].", UserName );
        }

        SqlQuery = L"USE master";
        m_AuditDatabase.ExecuteDirect( SqlQuery );

        SqlQuery = L"CREATE DATABASE ";
        SqlQuery += m_Settings.DatabaseName;
        m_AuditDatabase.ExecuteDirect( SqlQuery );

        SqlQuery = L"USE ";
        SqlQuery += m_Settings.DatabaseName;
        m_AuditDatabase.ExecuteDirect( SqlQuery );
    }

    // Tables
    AddCreateComputerMasterSql( AutoIncrementKeyword,
                                IntegerKeyword,
                                TimestampKeyword, VarCharKeyword, VarTCharKeyword, &Statements );
    AddCreateAuditMasterSql( AutoIncrementKeyword,
                             IntegerKeyword,
                             TimestampKeyword, VarCharKeyword, VarTCharKeyword, &Statements );
    AddCreateAuditDataSql( IntegerKeyword, TCharKeyword, VarTCharKeyword, &Statements );
    AddCreateDisplayNamesSql( IntegerKeyword, VarTCharKeyword, &Statements );

    // Views
    AddCreateViews( SchemaName, &Statements );

    // Procedures and Grants
    if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_MYSQL ) == 0 )
    {
        AddCreateMySqlProcedures( &Statements );
    }
    else if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_POSTGRE_SQL ) == 0 )
    {
        // No procedures for PostgreSQL
        if ( m_GrantPublicCheckBox.GetState() )
        {
            AddCreateGrants( &Statements );
        }
    }
    else if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_SQL_SERVER ) == 0 )
    {
        AddCreateSqlServerProcedures( &Statements );
        if ( m_GrantPublicCheckBox.GetState() )
        {
            AddCreateGrants( &Statements );
        }
    }

    // Execute the statements;
    WaitCursor Cursor;
    try
    {
        m_ProgressBar.SetPercentage( 0 );
        m_ProgressBar.SetVisible( true );
        m_ProgressBar.RedrawNow();
        m_AuditDatabase.BeginTrans();
        numStatements = Statements.GetSize();
        for ( i = 0; i < numStatements; i++ )
        {
            SqlQuery = Statements.Get( i );
            m_AuditDatabase.ExecuteTrans( SqlQuery );
            percentage = PXSCastSizeTToUInt32( ( i * 100 ) / numStatements );
            m_ProgressBar.SetPercentage( percentage );
        }
        m_AuditDatabase.CommitTrans();
        m_ProgressBar.SetPercentage( 100 );
    }
    catch ( const Exception& )
    {
        m_ProgressBar.SetVisible( false );
        throw;
    }
    m_ProgressBar.SetVisible( false );
    PXSGetResourceString( PXS_IDS_1238_DATABASE_WAS_CREATED, &ProgressMessage );
    SetProgressMessage( ProgressMessage );
}

//===============================================================================================//
//  Description:
//      Delete the old audits, i.e. preserve the most recent audit for each
//      computer in the database
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::DeleteOldAudits()
{
    DWORD  percentage = 0;
    size_t i = 0, numStatements = 0;
    String ApplicationName, SqlQuery, SqlDelete, ComputerID, MaxOfAuditID;
    String Text, ProgressMessage;
    Formatter     Format;
    OdbcRecordSet RecordSet;
    StringArray DeleteStatements;

    // Make sure the controls have been created
    if ( m_bControlsCreated == false ) return;

    ProgressMessage = PXS_STRING_EMPTY;
    SetProgressMessage( ProgressMessage );
    PXSGetApplicationName( &ApplicationName );
    PXSGetResourceString( PXS_IDS_1240_DELETE_OLD_AUDITS, &Text );
    if ( IDYES != MessageBox( m_hWindow,
                              Text.c_str(), ApplicationName.c_str(), MB_YESNO | MB_ICONQUESTION ) )
    {
        return;
    }
    PXSLogAppInfo( L"Deleteing old audits." );

    // Get the maximum audit ID for each computer that has more than one audit
    // in the database
    ConnectDB( false );
    WaitCursor Wait;
    SqlQuery  = L"SELECT Audit_Master.Computer_ID, "
                L"MAX(Audit_Master.Audit_ID) AS MaxOfAudit_ID "
                L"FROM Audit_Master "
                L"GROUP BY Audit_Master.Computer_ID "
                L"HAVING ( ( ( COUNT(Audit_Master.Audit_ID) )>1 ) );";
    RecordSet.Open( SqlQuery, &m_AuditDatabase );
    while ( RecordSet.MoveNext() )
    {
        ComputerID   = RecordSet.FieldValue( 0 );
        MaxOfAuditID = RecordSet.FieldValue( 1 );
        if ( ComputerID.GetLength() && MaxOfAuditID.GetLength() )
        {
            SqlDelete  = L"Delete FROM Audit_Master WHERE Computer_ID=";
            SqlDelete += ComputerID;
            SqlDelete += L" AND Audit_ID< ";
            SqlDelete += MaxOfAuditID;
            DeleteStatements.Add( SqlDelete );
        }
    }
    RecordSet.Close();

    // Is there anything to do?
    numStatements = DeleteStatements.GetSize();
    PXSLogAppInfo1( L"Found old audits for %%1 computers.", Format.SizeT(  numStatements ) );
    if ( numStatements == 0 )
    {
        PXSGetResourceString( PXS_IDS_1241_NO_OLD_AUDITS, &Text );
        MessageBox( m_hWindow, Text.c_str(), ApplicationName.c_str(), MB_OK | MB_ICONINFORMATION );
        return;
    }

    // Do the deletes
    m_AuditDatabase.BeginTrans();
    try
    {
        m_ProgressBar.SetPercentage( 0 );
        m_ProgressBar.SetVisible( true );
        for ( i = 0; i < numStatements; i++ )
        {
             percentage = PXSCastSizeTToUInt32( ( i * 100 ) / numStatements );
            m_ProgressBar.SetPercentage( percentage );
            SqlDelete = DeleteStatements.Get( i );
            m_AuditDatabase.ExecuteTrans( SqlDelete, PXS_SQLLEN_MAX );
        }
    }
    catch ( const Exception& )
    {
        m_ProgressBar.SetVisible( false );
        m_AuditDatabase.RollbackTrans();
        throw;
    }
    m_AuditDatabase.CommitTrans();
    m_ProgressBar.SetVisible( false );
    PXSGetResourceString( PXS_IDS_1242_DELETED_OLD_AUDITS, &ProgressMessage );
    SetProgressMessage( ProgressMessage );
}

//===============================================================================================//
//  Description:
//      Get the SQL keywords used by the database
//
//  Parameters:
//      pAutoIncrementKeyword - the database specific auto-increment keyword
//      pCharKeyword          - the database specific SQL_CHAR keyword
//      pIntegerKeyword       - the database specific SQL_INTEGER keyword
//      pSmallIntKeyword      - the database specific SQL_SMALLINT keyword
//      pTimestampKeyword     - the database specific SQL_TYPE_TIMESTAMP keyword
//      pVarCharKeyword       - the database specific SQL_VARCHAR keyword
//      pWCharKeyword         - the database specific SQL_WCHAR keyword
//      pWVarTCharKeyword     - the database specific SQL_WVARCHAR keyword
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::GetDatabaseSpecificKeywords(
                                                  String* pAutoIncrementKeyword,
                                                  String* pCharKeyword,
                                                  String* pIntegerKeyword,
                                                  String* pSmallIntKeyword,
                                                  String* pTimestampKeyword,
                                                  String* pVarCharKeyword,
                                                  String* pWCharKeyword,
                                                  String* pWVarCharKeyword )
{
    String    Message;
    Formatter Format;

    if ( ( pAutoIncrementKeyword == nullptr ) ||
         ( pCharKeyword          == nullptr ) ||
         ( pIntegerKeyword       == nullptr ) ||
         ( pSmallIntKeyword      == nullptr ) ||
         ( pTimestampKeyword     == nullptr ) ||
         ( pVarCharKeyword       == nullptr ) ||
         ( pWCharKeyword         == nullptr ) ||
         ( pWVarCharKeyword      == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }

    if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_ACCESS ) == 0 )
    {
        *pAutoIncrementKeyword = L"AUTOINCREMENT";
        *pCharKeyword          = L"VARCHAR";     // 12
        *pIntegerKeyword       = L"INTEGER";     // 4
        *pSmallIntKeyword      = L"SHORT";       // 5
        *pTimestampKeyword     = L"DATETIME";    // -93
        *pVarCharKeyword       = L"VARCHAR";     // 12
        *pWCharKeyword         = L"NVARCHAR";    // -8
        *pWVarCharKeyword      = L"NVARCHAR";    // -9
    }
    else
    {
        // Server database
        m_AuditDatabase.GetSupportsSqlType( SQL_CHAR          ,  pCharKeyword );
        m_AuditDatabase.GetSupportsSqlType( SQL_INTEGER       , pIntegerKeyword );
        m_AuditDatabase.GetSupportsSqlType( SQL_SMALLINT      , pSmallIntKeyword );
        m_AuditDatabase.GetSupportsSqlType( SQL_TYPE_TIMESTAMP, pTimestampKeyword );
        m_AuditDatabase.GetSupportsSqlType( SQL_VARCHAR       , pVarCharKeyword );
        m_AuditDatabase.GetSupportsSqlType( SQL_WCHAR         , pWCharKeyword );
        m_AuditDatabase.GetSupportsSqlType( SQL_WVARCHAR      , pWVarCharKeyword );

        if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_MYSQL ) == 0 )
        {
            // Will use 32-bit INT even though LAST_INSERT_ID returns 64-bit
            *pAutoIncrementKeyword = L"INT AUTO_INCREMENT";
        }
        else if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_SQL_SERVER ) == 0 )
        {
            *pAutoIncrementKeyword = L"int IDENTITY";    // need int

            // SQLGetTypeInfo when used with SQL Server 2008 + Native
            // Client 10.0 says SQL_WVARCHAR is "date". Not clear why that is.
            if ( pWVarCharKeyword->CompareI( L"nvarchar" ) )
            {
                PXSLogAppInfo1( L"Overriding '%%1' with nvarchar.", *pWVarCharKeyword );
                *pWVarCharKeyword = L"nvarchar";
            }
        }
        else if ( m_Settings.DBMS.CompareI( PXS_DBMS_NAME_POSTGRE_SQL ) == 0 )
        {
            *pAutoIncrementKeyword = L"SERIAL";
        }
        else
        {
            // Unknown DBMS
            Message = Format.String1( L"m_Settings.DBMS = '%%1'", m_Settings.DBMS );
            throw SystemException( ERROR_INVALID_NAME,
                                   Message.c_str(), __FUNCTION__ );
        }
    }

    PXSLogAppInfo1( L"AUTOINCREMENT     : '%%1'", *pAutoIncrementKeyword );
    PXSLogAppInfo1( L"SQL_CHAR          : '%%1'", *pCharKeyword );
    PXSLogAppInfo1( L"SQL_INTEGER       : '%%1'", *pIntegerKeyword );
    PXSLogAppInfo1( L"SQL_SMALLINT      : '%%1'", *pSmallIntKeyword );
    PXSLogAppInfo1( L"SQL_TYPE_TIMESTAMP: '%%1'", *pTimestampKeyword );
    PXSLogAppInfo1( L"SQL_VARCHAR       : '%%1'", *pVarCharKeyword );
    PXSLogAppInfo1( L"SQL_WCHAR         : '%%1'", *pWCharKeyword );
    PXSLogAppInfo1( L"AQL_WVARCHAR      : '%%1'", *pWVarCharKeyword );
}

//===============================================================================================//
//  Description:
//      Make the sql statement for the specified report
//
//  Parameters:
//      UpperKeyword - the database specific UPPER keyword, can be empty
//      pSqlQuery    - receives the  SQL statement
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::MakeReportSql( const String& UpperKeyword, String* pSqlQuery )
{
    String Sql;
    Formatter Format;

    if ( pSqlQuery == nullptr )
    {
        throw ParameterException( L"pSqlQuery", __FUNCTION__ );
    }
    *pSqlQuery = PXS_STRING_EMPTY;

    switch( m_ReportNamesComboBox.GetSelectedIndex() )
    {
        // Fall through
        default:
        case 0:
            // Automatic Updates Status
            Sql = L"SELECT Computer_Master.Domain_Name, "
                  L"Computer_Master.Computer_Name, Audit_Data.Item_3 AS "
                  L"Update_Status \r\n"
                  L"FROM Computer_Master INNER JOIN Audit_Data ON "
                  L"Computer_Master.Last_Audit_ID = Audit_Data.Audit_ID \r\n"
                  L"WHERE ( ( Audit_Data.Category_ID = 1200 ) AND "
                  L"( %%1( Audit_Data.Item_1 ) = 'AUTOMATIC UPDATES' ) AND "
                  L"( %%1( Audit_Data.Item_2 ) = 'UPDATE STATUS' ) ) \r\n"
                  L"ORDER BY %%1( Computer_Master.Computer_Name ), "
                  L"Audit_Data.Category_ID";
            break;

        case 1:
            // Counts of Audits by Computer
            Sql = L"SELECT %%1( Computer_Master.Domain_Name ) AS "
                  L"Domain_Name, %%1( Computer_Master.Computer_Name ) AS "
                  L"Computer_Name, COUNT( Audit_Master.Audit_ID ) AS "
                  L"Audit_Count \r\n"
                  L"FROM Computer_Master \r\n"
                  L"INNER JOIN Audit_Master ON Computer_Master.Computer_ID="
                  L"Audit_Master.Computer_ID \r\n"
                  L"GROUP BY %%1( Computer_Master.Domain_Name ), %%1( "
                  L"Computer_Master.Computer_Name ) \r\n"
                  L"ORDER BY %%1( Computer_Master.Domain_Name), "
                  L"%%1( Computer_Master.Computer_Name )";
            break;

        case 2:
            // Counts of Computers by Domain/Workgroup
            Sql = L"SELECT %%1( Computer_Master.Domain_Name ) AS "
                  L"Domain_Name, COUNT( Computer_Master.Computer_Name ) AS "
                  L"Computer_Count \r\n"
                  L"FROM Computer_Master \r\n"
                  L"GROUP BY %%1( Computer_Master.Domain_Name ) \r\n"
                  L"ORDER BY %%1( Computer_Master.Domain_Name )";
            break;

        case 3:
            // Counts of Errors by Type
            Sql = L"SELECT %%1( Audit_Data.Item_3 ) As Error_Type, "
                  L"COUNT( Audit_Data.Item_3 ) AS Error_Count \r\n"
                  L"FROM Computer_Master INNER JOIN Audit_Data ON "
                  L"Computer_Master.Last_Audit_ID = Audit_Data.Audit_ID \r\n"
                  L"WHERE Audit_Data.Category_ID = 2200 \r\n"
                  L"GROUP BY %%1( Audit_Data.Item_3 ) \r\n"
                  L"ORDER BY %%1( Audit_Data.Item_3 )";
            break;

        case 4:
            // Counts of Installed Software Products
            Sql = L"SELECT %%1( Audit_Data.Item_1 ) As Software_Name, "
                  L"COUNT( Audit_Data.Item_1 ) AS Installed_Count \r\n"
                  L"FROM Computer_Master INNER JOIN Audit_Data ON "
                  L"Computer_Master.Last_Audit_ID = Audit_Data.Audit_ID \r\n"
                  L"WHERE Audit_Data.Category_ID = 500 \r\n"
                  L"GROUP BY %%1( Audit_Data.Item_1 ) \r\n"
                  L"ORDER BY %%1( Audit_Data.Item_1 )";
            break;

        case 5:
            // Counts of Memory
            // For MySQL, v5 is required for ORDER BY an aggregate expression
            Sql = L"SELECT Audit_Data.Item_1 AS Memory_MB, COUNT( "
                  L"Audit_Data.Item_1 ) AS Computer_Count \r\n"
                  L"FROM Computer_Master INNER JOIN Audit_Data ON "
                  L"Computer_Master.Last_Audit_ID = Audit_Data.Audit_ID \r\n"
                  L"WHERE Audit_Data.Category_ID = 3600 \r\n"
                  L"GROUP BY Audit_Data.Item_1 \r\n"
                  L"ORDER BY COUNT( Audit_Data.Item_1 )";
            break;

        case 6:
            // Counts of Operating Systems
            Sql = L"SELECT %%1( Audit_Data.Item_1 ) AS Operating_System, "
                  L"COUNT( Audit_Data.Item_1 ) AS Computer_Count \r\n"
                  L"FROM Computer_Master INNER JOIN Audit_Data ON "
                  L"Computer_Master.Last_Audit_ID = Audit_Data.Audit_ID \r\n"
                  L"WHERE Audit_Data.Category_ID = 700  \r\n"
                  L"GROUP BY %%1( Audit_Data.Item_1 ) \r\n"
                  L"ORDER BY %%1( Audit_Data.Item_1 )";
            break;

        case 7:
            // Counts of Running Services
            Sql = L"SELECT %%1( Audit_Data.Item_1 ) AS Service_Name, "
                  L"COUNT( Audit_Data.Item_1 ) AS Service_Count \r\n"
                  L"FROM Computer_Master INNER JOIN Audit_Data ON "
                  L"Computer_Master.Last_Audit_ID = Audit_Data.Audit_ID \r\n"
                  L"WHERE Audit_Data.Category_ID = 4100  \r\n"
                  L"GROUP BY %%1( Audit_Data.Item_1 ) \r\n"
                  L"ORDER BY %%1( Audit_Data.Item_1 )";
            break;

        case 8:
            // Hardware Summary
            Sql = L"SELECT Computer_Master.Computer_Name, Audit_Data.Item_12"
                  L" AS Processor, Audit_Data.Item_13 AS Memory_MB, "
                  L"Audit_Data.Item_14 AS Size_Bytes, Audit_Data.Item_15 AS "
                  L"Display \r\n"
                  L"FROM Computer_Master INNER JOIN Audit_Data ON "
                  L"Computer_Master.Last_Audit_ID = Audit_Data.Audit_ID \r\n"
                  L"WHERE Audit_Data.Category_ID = 300 \r\n"
                  L"ORDER BY %%1( Computer_Master.Computer_Name )";
            break;

        case 9:
            // Last Audit Date by Computer
            Sql = L"SELECT Computer_Master.Domain_Name, "
                  L"Computer_Master.Computer_Name, "
                  L"Audit_Master.Database_Local AS Last_Audit_Date \r\n"
                  L"FROM Computer_Master \r\n"
                  L"INNER JOIN Audit_Master ON "
                  L"(Computer_Master.Last_Audit_ID = Audit_Master.Audit_ID)"
                  L" \r\n"
                  L"AND (Computer_Master.Computer_ID = "
                  L"Audit_Master.Computer_ID) \r\n"
                  L"ORDER BY Audit_Master.Database_Local DESC\r\n";
            break;

        case 10:
            // List BIOS Versions
            Sql = L"SELECT Computer_Master.Computer_Name, Audit_Data.Item_1 "
                  L"AS BIOS_Version, Audit_Data.Item_2 AS Release_Date \r\n"
                  L"FROM Computer_Master INNER JOIN Audit_Data ON "
                  L"Computer_Master.Last_Audit_ID = Audit_Data.Audit_ID \r\n"
                  L"WHERE Audit_Data.Category_ID = 3100 \r\n"
                  L"ORDER BY %%1( Computer_Master.Computer_Name )";
            break;

        case 11:
            // List Displays
            Sql = L"SELECT Computer_Master.Computer_Name, Audit_Data.Item_2 "
                  L"AS Display_Name, Audit_Data.Item_3 AS Manufacturer \r\n"
                  L"FROM Computer_Master INNER JOIN Audit_Data ON "
                  L"Computer_Master.Last_Audit_ID = Audit_Data.Audit_ID \r\n"
                  L"WHERE Audit_Data.Category_ID = 4600 \r\n"
                  L"ORDER BY %%1( Computer_Master.Computer_Name ), %%1( "
                  L"Audit_Data.Item_2 )";
            break;

        case 12:
            // List Installed Printers
            Sql = L"SELECT Computer_Master.Computer_Name, Audit_Data.Item_2 "
                  L"AS Printer_Name, Audit_Data.Item_3 AS Share_Name \r\n"
                  L"FROM Computer_Master INNER JOIN Audit_Data ON "
                  L"Computer_Master.Last_Audit_ID = Audit_Data.Audit_ID \r\n"
                  L"WHERE Audit_Data.Category_ID = 3000 \r\n"
                  L"ORDER BY %%1( Computer_Master.Computer_Name ), %%1( "
                  L"Audit_Data.Item_2 )";
            break;

        case 13:
            // List Network Adapters
            Sql = L"SELECT Computer_Master.Computer_Name, Audit_Data.Item_2 "
                  L"AS Adapter_Name, Audit_Data.Item_5 AS IP_Address, "
                  L"Audit_Data.Item_18 AS Connection_Speed_MBPS\r\n"
                  L"FROM Computer_Master INNER JOIN Audit_Data ON "
                  L"Computer_Master.Last_Audit_ID = Audit_Data.Audit_ID \r\n"
                  L"WHERE Audit_Data.Category_ID = 2600 \r\n"
                  L"ORDER BY %%1( Computer_Master.Computer_Name ), %%1( "
                  L"Audit_Data.Item_2 )";
            break;

        case 14:
            // List Physical Disks
            Sql = L"SELECT Computer_Master.Computer_Name, Audit_Data.Item_5 "
                  L"AS Model, Audit_Data.Item_2 AS Capacity_MB\r\n"
                  L"FROM Computer_Master INNER JOIN Audit_Data ON "
                  L"Computer_Master.Last_Audit_ID = Audit_Data.Audit_ID \r\n"
                  L"WHERE Audit_Data.Category_ID = 3700 \r\n"
                  L"ORDER BY %%1( Computer_Master.Computer_Name ), %%1( "
                  L"Audit_Data.Item_5 )";
            break;

        case 15:
            // List Processors
            Sql = L"SELECT Computer_Master.Computer_Name, Audit_Data.Item_2 "
                  L"AS Processor_Name, Audit_Data.Item_4 AS "
                  L"Estimated_Speed_MHz \r\n"
                  L"FROM Computer_Master INNER JOIN Audit_Data ON "
                  L"Computer_Master.Last_Audit_ID = Audit_Data.Audit_ID \r\n"
                  L"WHERE Audit_Data.Category_ID = 3200 \r\n"
                  L"ORDER BY %%1( Computer_Master.Computer_Name )";
            break;

        case 16:
            // List Users
            Sql = L"SELECT Computer_Master.Domain_Name, Audit_Data.Item_1 "
                  L"AS User_Name, Audit_Data.Item_4 AS Account_Status, "
                  L"Audit_Data.Item_9 AS Number_Of_Logons \r\n"
                  L"FROM Computer_Master INNER JOIN Audit_Data ON "
                  L"Computer_Master.Last_Audit_ID = Audit_Data.Audit_ID \r\n"
                  L"WHERE Audit_Data.Category_ID = 1900 \r\n"
                  L"ORDER BY %%1( Computer_Master.Domain_Name ), "
                  L"%%1( Audit_Data.Item_1 )";
            break;
    }
    Sql.ReplaceI( L"%%1", UpperKeyword.c_str() );
    *pSqlQuery = Sql;
}

//===============================================================================================//
//  Description:
//      Run the selected report
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::RunReport()
{
    const size_t TABLE_WIDTH = 9000;
    DWORD  maxRecords = 0, recordCount = 0;
    size_t i = 0, columnCount = 0, columnWidth = 0;
    String UpperKeyword, SqlQuery, ReportName, RichText, UserName, Text;
    String ApplicationName, Separator;
    String HeaderCol_1, HeaderCol_2, RowEven_1, RowEven_2, RowOdd_1, RowOdd_2;
    Formatter      Format;
    OdbcRecordSet  RecordSet;
    MessageDialog  ReportDialog;
    SystemInformation SystemInfo;

    // Markup strings
    LPCWSTR STR_SEPARATOR    = L"\\par\\trowd\\trgaph108\\trleft-108"
                               L"\\trrh-60\\clcbpat1\\cellx%%1\\pard\\intbl"
                               L"\\cell\\row\\pard\\par\r\n";
    LPCWSTR STR_HEADER_COL_1 = L"\\clcbpat1\\clbrdrl\\brdrw10\\brdrs"
                               L"\\clbrdrt\\brdrw10\\brdrs\\clbrdrr\\brdrw10"
                               L"\\brdrs\\clbrdrb\\brdrw10\\brdrs"
                               L"\\cellx%%1\r\n";
    LPCWSTR STR_HEADER_COL_2 = L"\\clcbpat1\\clbrdrl\\brdrw10\\brdrs"
                               L"\\clbrdrt\\brdrw10\\brdrs\\clbrdrr\\brdrw10"
                               L"\\brdrs\\clbrdrb\\brdrw10\\brdrs"
                               L"\\cellx%%1\r\n";
    LPCWSTR STR_ROW_EVEN_1   = L"\\clcbpat2\\clbrdrl\\brdrw10\\brdrs"
                               L"\\clbrdrt\\brdrw10\\brdrs\\clbrdrr\\brdrw10"
                               L"\\brdrs\\clbrdrb\\brdrw10\\brdrs"
                               L"\\cellx%%1\r\n";
    LPCWSTR STR_ROW_EVEN_2   = L"\\clcbpat2\\clbrdrl\\brdrw10\\brdrs"
                               L"\\clbrdrt\\brdrw10\\brdrs\\clbrdrr\\brdrw10"
                               L"\\brdrs\\clbrdrb\\brdrw10\\brdrs"
                               L"\\cellx%%1\r\n";
    LPCWSTR STR_ROW_ODD_1    = L"\\clbrdrl\\brdrw10\\brdrs\\clbrdrt\\brdrw10"
                               L"\\brdrs\\clbrdrr\\brdrw10\\brdrs\\clbrdrb"
                               L"\\brdrw10\\brdrs \\cellx%%1\r\n";
    LPCWSTR STR_ROW_ODD_2    = L"\\clbrdrl\\brdrw10\\brdrs\\clbrdrt\\brdrw10"
                               L"\\brdrs\\clbrdrr\\brdrw10\\brdrs\\clbrdrb"
                               L"\\brdrw10\\brdrs \\cellx%%1\r\n";

    Separator   = Format.StringUInt32( STR_SEPARATOR   , TABLE_WIDTH );
    HeaderCol_1 = Format.StringUInt32( STR_HEADER_COL_1, TABLE_WIDTH / 3 );
    HeaderCol_2 = Format.StringUInt32( STR_HEADER_COL_2, TABLE_WIDTH );
    RowEven_1   = Format.StringUInt32( STR_ROW_EVEN_1  , TABLE_WIDTH / 3 );
    RowEven_2   = Format.StringUInt32( STR_ROW_EVEN_2  , TABLE_WIDTH );
    RowOdd_1    = Format.StringUInt32( STR_ROW_ODD_1   , TABLE_WIDTH / 3 );
    RowOdd_2    = Format.StringUInt32( STR_ROW_ODD_2   , TABLE_WIDTH );

    // Get the UPPER keyword if its required
    ConnectDB( false );
    UpperKeyword = PXS_STRING_EMPTY;
    if ( m_AuditDatabase.IsCaseSensitiveSort() )
    {
       m_AuditDatabase.GetDbKeyWord( PXS_KEYWORD_UPPER, &UpperKeyword );
    }

    MakeReportSql( UpperKeyword, &SqlQuery );
    maxRecords = m_MaxReportRowsSpinner.GetValue();
    if ( maxRecords == 0 )      // if its zero, fetch them all
    {
        maxRecords = UINT32_MAX;
    }

    WaitCursor Wait;
    RichText.Allocate( 4096 );
    PXSGetRichTextDocumentStart( &RichText );
    RecordSet.Open( SqlQuery, &m_AuditDatabase );
    columnCount = RecordSet.GetColumnCount();
    if ( columnCount )
    {
        columnWidth = ( TABLE_WIDTH / columnCount );
    }

    // Title
    m_ReportNamesComboBox.GetSelectedString( &Text );
    RichText += L"\\par\\qc\\ul\\b ";
    RichText += Text;
    RichText += L" \\b0\\ul0\\par\\par\\pard\r\n";

    // Table header
    RichText += L"\\trowd\\trgaph108\\trleft-108\r\n";
    for ( i = 0; i < columnCount; i++ )
    {
        RichText += L"\\clcbpat1\\clbrdrl\\brdrw10\\brdrs\\clbrdrt\\brdrw10"
                    L"\\brdrs\\clbrdrr\\brdrw10\\brdrs\\clbrdrb\\brdrw10\\"
                    L"brdrs\\cellx";
        RichText += Format.SizeT( ( i + 1 ) * columnWidth );
        RichText += PXS_STRING_CRLF;
    }
    RichText += L"\\pard\\intbl\r\n";
    for ( i = 0; i < columnCount; i++ )
    {
        Text = RecordSet.GetColumnDisplayName( i );
        Text.EscapeForRichText();
        RichText += L"\\b ";
        RichText += Text;
        RichText += L"\\b0\\cell\r\n";
    }
    RichText += L"\\row\r\n\r\n";

    // Rows.
    while ( ( recordCount < maxRecords ) && RecordSet.MoveNext() )
    {
        RichText += L"\\trowd\\trgaph108\\trleft-108\r\n";
        for ( i = 0; i < columnCount; i++ )
        {
            if ( recordCount % 2 )
            {
                RichText += L"\\clcbpat2";
            }
            RichText += L"\\clbrdrl\\brdrw10\\brdrs\\clbrdrt\\brdrw10"
                        L"\\brdrs\\clbrdrr\\brdrw10\\brdrs\\clbrdrb\\"
                        L"brdrw10\\brdrs\\cellx";
            RichText += Format.SizeT( ( i + 1 ) * columnWidth );
            RichText += PXS_STRING_CRLF;
        }
        RichText += L"\\pard\\intbl\r\n";
        for ( i = 0; i < columnCount; i++ )
        {
            Text = RecordSet.FieldValue( i );
            Text.EscapeForRichText();
            RichText += Text;
            RichText += L" \\cell\r\n";
        }
        RichText += L"\\row\r\n\r\n";
        recordCount++;
    }
    RecordSet.Close();

    // Table Close
    RichText += L"\\pard\r\n\r\n";
    RichText += Separator;

    // Add the user name, timestamp and row count
    SystemInfo.GetCurrentUserName( &UserName );
    RichText += L"User:\t";
    RichText += UserName;
    RichText += L"\\par ";

    RichText += L"Date:\t";
    RichText += Format.LocalTimeInUserLocale();
    RichText += L"\\par ";

    RichText += L"Rows:\t";
    RichText += Format.UInt32( recordCount );
    RichText += L"\\par ";

    // Add the SQL statement
    if ( m_ReportShowSqlCheckBox.GetState() )
    {
        RichText += L"SQL:\r\n";
        RichText += L"\\par ";

        Text = SqlQuery;
        Text.EscapeForRichText();
        RichText += L"\\trowd\\trgaph108\\trleft-108\r\n";
        RichText += L"\\clbrdrl\\brdrw10\\brdrs\\clbrdrt\\brdrw10\\brdrs"
                    L"\\clbrdrr\\brdrw10\\brdrs\\clbrdrb\\brdrw10\\brdrs"
                    L"\\cellx";
        RichText += Format.UInt32( TABLE_WIDTH );
        RichText += PXS_STRING_CRLF;
        RichText += L"\\pard\\intbl\r\n";
        RichText += Text;
        RichText += L"\\cell\r\n";     // Close cell
        RichText += L"\\row\r\n";      // Close row
        RichText += L"\\pard\\r\n";    // Close table
    }

    // Footer
    RichText += Separator;
    PXSGetApplicationName( &ApplicationName );
    ApplicationName.EscapeForRichText();
    RichText += L"\\qc ";
    RichText += L"Generated by ";
    RichText += ApplicationName;
    RichText += L" \\par\r\n";

    // Document end
    RichText += L"}";

    // Show it, set the text then assign the dynamic chart
    ReportDialog.SetTitle( ApplicationName );
    ReportDialog.SetSize( 650, 475 );
    ReportDialog.SetMessage( RichText );
    ReportDialog.Create( m_hWindow );
}

//===============================================================================================//
//  Description:
//      Put the control values in the configuration
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::UpdateSettings()
{
    m_ReportNamesComboBox.GetSelectedString( &m_Settings.LastReportName );
    m_Settings.reportMaxRecords = m_MaxReportRowsSpinner.GetValue();
    m_Settings.reportShowSQL    = m_ReportShowSqlCheckBox.GetState();
}

//===============================================================================================//
//  Description:
//      Set text progress message on the dialog box
//
//  Parameters:
//      ProgressMessage - 1-line message
//
//   Returns:
//      void
//===============================================================================================//
void DatabaseAdministrationDialog::SetProgressMessage(
                                                const String& ProgressMessage )
{
    RECT  bounds = { 0, 0, 0, 0 };
    StaticControl Static;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    if ( m_uIdxStaticProgress < m_Statics.GetSize() )
    {
        Static = m_Statics.Get( m_uIdxStaticProgress );
        Static.SetText( ProgressMessage );
        m_Statics.Set( m_uIdxStaticProgress, Static );
        Static.GetBounds( &bounds );
        InvalidateRect( m_hWindow, &bounds, TRUE );
        RedrawWindow( m_hWindow, &bounds, nullptr, RDW_UPDATENOW );
    }
}
