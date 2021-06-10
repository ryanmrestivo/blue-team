///////////////////////////////////////////////////////////////////////////////////////////////////
//
// WinAudit Frame Class Implementation
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

// Must have allocated the g_pApplication object

///////////////////////////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/WinAuditFrame.h"

// 2. C System Files
#include <time.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/AboutDialog.h"
#include "PxsBase/Header Files/AllocateChars.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/AutoUnlockMutex.h"
#include "PxsBase/Header Files/ComException.h"
#include "PxsBase/Header Files/Directory.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/Mail.h"
#include "PxsBase/Header Files/MessageDialog.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/Shell.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/WaitCursor.h"

// 5. This Project
#include "WinAudit/Header Files/AuditData.h"
#include "WinAudit/Header Files/CpuInformation.h"
#include "WinAudit/Header Files/DiskInformation.h"
#include "WinAudit/Header Files/DisplayInformation.h"
#include "WinAudit/Header Files/OdbcExportDialog.h"
#include "WinAudit/Header Files/Resources.h"
#include "WinAudit/Header Files/SecurityInformation.h"
#include "WinAudit/Header Files/SoftwareInformation.h"
#include "WinAudit/Header Files/TcpIpInformation.h"
#include "WinAudit/Header Files/WinAuditConfigDialog.h"
#include "WinAudit/Header Files/WindowsInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
WinAuditFrame::WinAuditFrame()
              :m_bAuditing( false ),
               m_bCreatedReport( false ),
               m_bSavedReport( false ),
               m_bCreatedControls( false ),
               m_uSaveAuditFilterIndex( 0 ),
               m_uTableCounter( 0 ),
               TAB_INDEX_AUDIT( 0 ),
               TAB_INDEX_PROCESSORS( 4 ),
               TAB_INDEX_DISKS( 1 ),
               TAB_INDEX_DISPLAYS( 2 ),
               TAB_INDEX_FIRMWARE( 3 ),
               TAB_INDEX_SOFTWARE( 5 ),
               TAB_INDEX_HELP( 6 ),
               TAB_INDEX_LOGGER( 7 ),
               DATA_FORMAT_CSV( 1 ),
               DATA_FORMAT_CSV2( 2 ),
               DATA_FORMAT_RTF( 3 ),
               DATA_FORMAT_HTML( 4 ),
               DATA_FORMAT_DEFAULT( DATA_FORMAT_HTML ),
               m_AuditThread(),
               m_ConfigurationSettings(),
               m_AuditRecords(),
               m_bRunAuditThreadMT( FALSE ),
               m_uAuditPercentDoneMT( 0 ),
               m_AuditThreadParameterMT(),
               m_AuditThreadExceptionsMT(),
               m_NewAuditRecordsMT(),
               m_Mutex(),
               m_MenuBar(),
               m_FileMenuBarItem(),
               m_FilePopup(),
               m_FileAudit(),
               m_FileSave(),
               m_FileEmail(),
               m_FileOdbc(),
               m_FileShortcut(),
               m_FileExit(),
               m_EditMenuBarItem(),
               m_EditPopup(),
               m_EditSelectAll(),
               m_EditCopy(),
               m_EditFind(),
               m_EditSetFont(),
               m_EditBiggerText(),
               m_EditSmallerText(),
               m_ViewMenuBarItem(),
               m_ViewPopup(),
               m_ViewToolBar(),
               m_ViewStatusBar(),
               m_ViewOptions(),
               m_ViewDiskInfo(),
               m_ViewDisplayInfo(),
               m_ViewFirmwareInfo(),
               m_ViewPolicyInfo(),
               m_ViewProcessorInfo(),
               m_ViewSoftwareInfo(),
               m_LanguageMenuBarItem(),
               m_LanguagePopup(),
               m_LanguageChineseTw(),
               m_LanguageCzech(),
               m_LanguageDanish(),
               m_LanguageGerman(),
               m_LanguageEnglish(),
               m_LanguageSpanish(),
               m_LanguageGreek(),
               m_LanguageBelgian(),
               m_LanguageFrench(),
               m_LanguageHebrew(),
               m_LanguageItalian(),
               m_LanguageIndonesian(),
               m_LanguageJapanese(),
               m_LanguageKorean(),
               m_LanguageHungarian(),
               m_LanguageDutch(),
               m_LanguagePolish(),
               m_LanguageBrazilian(),
               m_LanguagePortuguese(),
               m_LanguageRussian(),
               m_LanguageSlovak(),
               m_LanguageSerbian(),
               m_LanguageFinnish(),
               m_LanguageThai(),
               m_LanguageTurkish(),
               m_HelpMenuBarItem(),
               m_HelpPopup(),
               m_HelpUsingWinAudit(),
               m_HelpWebsite(),
               m_HelpAbout(),
               m_HelpCreateGuid(),
               m_HelpLogFile(),
               m_MainPanel(),
               m_ToolBar(),
               m_AuditButton(),
               m_StopButton(),
               m_OptionsButton(),
               m_SaveButton(),
               m_EmailButton(),
               m_HelpButton(),
               m_WorkSpace(),
               m_FindTextBar(),
               m_Splitter(),
               m_SidePanel(),
               m_AuditCategories(),
               m_HelpContents(),
               m_TabWindow(),
               m_AuditRichBox(),
               m_DiskRichBox(),
               m_DisplayRichBox(),
               m_FirmwareRichBox(),
               m_ProcessorRichBox(),
               m_SoftwareRichBox(),
               m_HelpRichBox(),
               m_AboutRichBox(),
               m_LogRichBox(),
               m_StatusBar(),
               m_ProgressBar()
{
}

// Copy constructor - not allowed so no implementation

// Destructor
WinAuditFrame::~WinAuditFrame()
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
//      Callback/sink for the audit thread to post new data
//
//  Parameters:
//      timeRemaining - time remaining before job time out
//      percentDone   - percentage complete
//      pAuditRecords - new data records
//      pAuditThreadException - possible exception raised in worker thread
//
//  Remarks:
//      This method is only called by the audit thread
//
//  Returns:
//      Error code. Non-zero tells the thread to stop working
//===============================================================================================//
DWORD WinAuditFrame::AuditThreadCallback( DWORD timeRemaining,
                                          DWORD percentDone,
                                          const TArray< AuditRecord >* pAuditRecords,
                                          const Exception* pAuditThreadException )
{
    // Test if need to stop
    if ( m_bRunAuditThreadMT == FALSE )
    {
        return ERROR_CANCELLED;
    }
    m_Mutex.TryLock( timeRemaining );
    AutoUnlockMutex AutoUnlock( &m_Mutex );

    // Store the new data
    m_uAuditPercentDoneMT = percentDone;
    if ( pAuditRecords )
    {
        m_NewAuditRecordsMT.Append( *pAuditRecords );
    }

    if ( pAuditThreadException )
    {
        m_AuditThreadExceptionsMT.Add( *pAuditThreadException );
    }

    // Tell the application there is new data to process
    PostMessage( m_hWindow, PXS_APP_MSG_AUDIT_THREAD_UPDATE, 0, 0 );

    return ERROR_SUCCESS;   // = Continue running
}

//===============================================================================================//
//  Description:
//      Show the result of the audit
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::AuditThreadDone()
{
    String ApplicationName, RichText;

    // Set the application's state
    m_bAuditing = false;

    // Show the result
    ShowHideSidePanel( true );
    m_AuditCategories.SetVisible( true );
    m_ProgressBar.SetPercentage( 100 );
    m_HelpContents.SetVisible( false );
    m_TabWindow.SetSelectedTabIndex( TAB_INDEX_AUDIT );
    m_TabWindow.SetTabVisible( TAB_INDEX_AUDIT, true );
    SetControlsAndMenus();
    DoLayout();

    // Set the document end
    PXSGetApplicationName( &ApplicationName );
    PXSGetRichTextDocumentStart( &RichText );
    RichText += L"\\qc\\cf3 Generated by ";
    RichText += ApplicationName;
    RichText += L" \\cf0\\par\\par\\pard\r\n";
    RichText += L"}";
    m_AuditRichBox.AppendRichText( RichText );
    m_AuditRichBox.ScrollToTop();

    PXSLogAppInfo( L"Audit job finished." );
}

//===============================================================================================//
//  Description:
//      Start the computer audit
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::AuditStart()
{
    String    ComputerName, RichText, EmptyString;
    Formatter Format;
    SystemInformation SystemInfo;

    // Don't start an audit if one is in progress
    if ( m_bAuditing )
    {
        throw FunctionException( L"m_bAuditing", __FUNCTION__ );
    }
    WaitCursor Wait;

    // Remove any existing audit
    EmptyString = PXS_STRING_EMPTY;
    m_uTableCounter = 0;
    m_AuditRichBox.SetText( EmptyString );
    m_AuditCategories.ClearList();
    m_AuditRecords.RemoveAll();
    Repaint();

    // Document header
    PXSGetRichTextDocumentStart( &RichText );
    SystemInfo.GetComputerNetBiosName( &ComputerName );
    RichText += L"\\pard\\qc \\b \\ul Computer Audit";
    if ( ComputerName.GetLength() )
    {
        RichText += L" for ";
        RichText += ComputerName;
    }
    RichText += L" \\ul0\\b0\\par\\par\\pard\r\n";
    RichText += L"}";
    m_AuditRichBox.SetRichText( RichText );

    // Reset shared variables
    if ( g_pApplication )
    {
        g_pApplication->SetStopBackgroundTasks( false );
    }
    m_bRunAuditThreadMT   = TRUE;
    m_uAuditPercentDoneMT = 0;
    m_AuditThreadExceptionsMT.RemoveAll();
    m_AuditThreadParameterMT.Categories.RemoveAll();

    ///////////////////////////////////////////////////////////////////////////
    // Fill the audit thread's parameters

    // No time out as user can cancel in the UI
    m_AuditThreadParameterMT.timeoutAt      = PXS_TIME_MAX;
    m_AuditThreadParameterMT.pWinAuditFrame = this;

    // The computer local time, this ensures have the same value anywhere a
    // timestamp is required
    m_AuditThreadParameterMT.LocalTime = Format.LocalTimeInIsoFormat();

    // Set the categories of data to be gathered
    m_ConfigurationSettings.MakeDataCategoriesArray( &m_AuditThreadParameterMT.Categories );
    if ( m_AuditThreadParameterMT.Categories.GetSize() == 0 )
    {
        PXSLogAppInfo( L"No options were selected, there is nothing to do.");
        return;     // Nothing to do
    }

    ///////////////////////////////////////////////////////////////////////////
    // Start

    if ( m_AuditThread.IsCreated() == false )
    {
        m_AuditThread.Run( m_hWindow );
    }
    m_AuditThread.SetAuditThreadParameter( &m_AuditThreadParameterMT );
    PXSLogAppInfo( L"Audit started." );

    // Set the state of the application
    m_bAuditing      = true;
    m_bCreatedReport = false;
    m_bSavedReport   = false;
    m_ProgressBar.SetPercentage( 2 );       // Small value to indicate started
    m_AuditCategories.SetVisible( true );
    m_HelpContents.SetVisible( false );
    ShowHideSidePanel( true );
    m_TabWindow.SetSelectedTabIndex( TAB_INDEX_AUDIT );
    m_TabWindow.SetTabVisible( TAB_INDEX_AUDIT, true );

    SetControlsAndMenus();
    DoLayout();
}

//===============================================================================================//
//  Description:
//      Update the result as new data is posted by the audit thread
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::AuditThreadUpdate()
{
    DWORD     categoryID = 0, percentDone;
    size_t    i = 0, numElements = 0;
    String    RichText, CategoryName;
    Formatter Format;
    AuditData Auditor;
    Exception ThreadException;
    AuditRecord  Record;
    TreeViewItem CategoryItem;
    TArray< Exception >    ThreadExceptions;
    TArray< AuditRecord >  NewAuditRecords;
    TArray< TreeViewItem > CategoryItems;

    // Copy the shared data
    m_Mutex.Lock();
    AutoUnlockMutex AutoUnlock( &m_Mutex );

    percentDone      = m_uAuditPercentDoneMT;
    NewAuditRecords  = m_NewAuditRecordsMT;
    ThreadExceptions = m_AuditThreadExceptionsMT;

    // Reset for next for set of new data
    m_NewAuditRecordsMT.RemoveAll();
    m_AuditThreadExceptionsMT.RemoveAll();

    // Log error any errors reported by the worker
    numElements = ThreadExceptions.GetSize();
    for ( i = 0; i < numElements; i++ )
    {
        ThreadException = ThreadExceptions.Get( i );
        if ( ThreadException.GetErrorCode() )
        {
            PXSLogAppError( L"An error was reported by a worker thread. Details follow next." );
            PXSLogException( ThreadException, __FUNCTION__ );
        }
    }

    // Make the content
    numElements = NewAuditRecords.GetSize();
    if ( numElements )
    {
        PXSLogAppInfo1( L"Worker posted %%1 audit record(s).", Format.SizeT( numElements ) );
    }
    PXSAuditRecordsToContent( NewAuditRecords, &CategoryItems, &m_uTableCounter, &RichText );
    if ( RichText.GetLength() )
    {
        m_bCreatedReport = true;
    }

    // Set
    m_AuditRecords.Append( NewAuditRecords );
    m_ProgressBar.SetPercentage( percentDone );
    numElements = CategoryItems.GetSize();
    for ( i = 0; i < numElements; i++ )
    {
        CategoryItem = CategoryItems.Get( i );
        m_AuditCategories.AddItem( CategoryItem );
    }
    m_AuditRichBox.AppendRichText( RichText );

    if ( NewAuditRecords.GetSize() )
    {
        Record     = NewAuditRecords.Get( 0 );
        categoryID = Record.GetCategoryID();
        Auditor.GetCategoryName( categoryID, &CategoryName );
        m_ProgressBar.SetNote( CategoryName );
    }
}

//===============================================================================================//
//  Description:
//     Select the content of the selected tab window
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::CopySelection()
{
    HWND hWnd = GetFocus();
    if ( hWnd )
    {
        SendMessage( hWnd, WM_COPY, 0, 0 );
    }
}

//===============================================================================================//
//  Description:
//      Create the controls
//
//  Parameters:
//      None
//
//  Returns:
//     void
//===============================================================================================//
void WinAuditFrame::CreateControls()
{
    int    height     = 0;
    BYTE   panelIndex = 0;
    SIZE   size       = { 0, 0 };
    Font   CourierNewFont, VerdanaFont;
    String Caption, HelpRichText;
    SystemInformation SystemInfo;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    // One-shot
    if ( m_bCreatedControls )
    {
        throw FunctionException( L"m_bCreatedControls", __FUNCTION__ );
    }
    m_bCreatedControls = true;

    // Some windows require a fixed width font
    CourierNewFont.SetFaceName( L"Courier New" );
    CourierNewFont.SetPointSize( 10, nullptr );
    CourierNewFont.Create();

    // Keep this font in sync with PXSGetRichTextDocumentStart so that if the
    // user desires to set the font, the CHOOSEFONT dialogue will show the
    // same font as in the rich text's document header font table (\fonttbl)
    VerdanaFont.SetFaceName( L"Verdana" );
    VerdanaFont.SetPointSize( 9, nullptr );     // = \fs18
    VerdanaFont.Create();

    ///////////////////////////////////////////////////////////////////////////
    // Application Panel

    m_MainPanel.SetStyle( WS_CLIPCHILDREN, true );
    m_MainPanel.Create( this );

    ///////////////////////////////////////////////////////////////////////////
    // Tool Bar

    m_ToolBar.SetStyle( WS_BORDER, true );
    m_ToolBar.Create( &m_MainPanel );
    m_ToolBar.SetHideBitmaps( m_hWindow, IDB_CLOSE_16, IDB_CLOSE_ON_16 );
    m_ToolBar.SetDoubleBuffered( true );

    // Audit Button
    m_AuditButton.Create( &m_ToolBar );
    m_AuditButton.SetAppMessageListener( m_hWindow );
    m_AuditButton.SetAppMessageCode( PXS_APP_MSG_BUTTON_CLICK );

    // Stop Button
    m_StopButton.Create( &m_ToolBar );
    m_StopButton.SetAppMessageListener( m_hWindow );
    m_StopButton.SetAppMessageCode( PXS_APP_MSG_BUTTON_CLICK );

    // Options Button
    m_OptionsButton.Create( &m_ToolBar );
    m_OptionsButton.SetAppMessageListener( m_hWindow );
    m_OptionsButton.SetAppMessageCode( PXS_APP_MSG_BUTTON_CLICK );

    // Save Button
    m_SaveButton.Create( &m_ToolBar );
    m_SaveButton.SetAppMessageListener( m_hWindow );
    m_SaveButton.SetAppMessageCode( PXS_APP_MSG_BUTTON_CLICK );

    // Email Button
    m_EmailButton.Create( &m_ToolBar );
    m_EmailButton.SetAppMessageListener( m_hWindow );
    m_EmailButton.SetAppMessageCode( PXS_APP_MSG_BUTTON_CLICK );

    // Help Button
    m_HelpButton.Create( &m_ToolBar );
    m_HelpButton.SetAppMessageListener( m_hWindow );
    m_HelpButton.SetAppMessageCode( PXS_APP_MSG_BUTTON_CLICK );

    // Now that have added the controls to the toolbar set its height
    m_ToolBar.GetSize( &size );
    height = m_ToolBar.GetPreferredHeight();
    m_ToolBar.SetSize( size.cx, height );

    ///////////////////////////////////////////////////////////////////////////
    // Find Text Bar

    m_FindTextBar.Create( &m_MainPanel );
    m_FindTextBar.CreateControls( m_hWindow );
    m_FindTextBar.SetHideBitmaps( m_hWindow, IDB_CLOSE_16, IDB_CLOSE_ON_16 );
    m_FindTextBar.SetAppMessageListener( m_hWindow );
    m_FindTextBar.SetDoubleBuffered( true );
    m_FindTextBar.SetVisible( false );

    ///////////////////////////////////////////////////////////////////////////
    // Workspace Panel

    m_WorkSpace.SetStyle( WS_CLIPCHILDREN, true );
    m_WorkSpace.Create( &m_MainPanel );

    m_SidePanel.SetStyle( WS_CLIPCHILDREN, true );
    m_SidePanel.Create( &m_WorkSpace );

    // Categories window
    m_AuditCategories.Create( &m_SidePanel );
    m_AuditCategories.SetAppMessageListener( m_hWindow );
    m_AuditCategories.SetBitmaps( IDB_TREE_NODE_CLOSED,
                                  IDB_TREE_NODE_OPEN,
                                  IDB_TREE_LEAF,
                                  PXS_COLOUR_WHITE,
                                  0,
                                  0,
                                  0,
                                  0 );
    m_AuditCategories.SetHideBitmaps( m_hWindow, IDB_CLOSE_16, IDB_CLOSE_ON_16);
    m_AuditCategories.SetSize( 0, 100 );
    m_AuditCategories.SetDoubleBuffered( true );

    // Help Content, initially invisible
    m_HelpContents.Create( &m_SidePanel );
    m_HelpContents.SetAppMessageListener( m_hWindow );
    m_HelpContents.SetBitmaps( IDB_TREE_NODE_CLOSED,
                               IDB_TREE_NODE_OPEN,
                               IDB_TREE_LEAF,
                               PXS_COLOUR_WHITE,
                               0,
                               0,
                               0,
                               0 );
    m_HelpContents.SetHideBitmaps( m_hWindow, IDB_CLOSE_16, IDB_CLOSE_ON_16 );
    m_HelpContents.SetSize( 0, 100 );
    m_HelpContents.SetDoubleBuffered( true );
    m_HelpContents.SetVisible( false );
    FillHelpContents();

    m_TabWindow.Create( &m_WorkSpace );
    m_TabWindow.SetDoubleBuffered( true );
    m_TabWindow.SetAppMessageListener( m_hWindow );
    m_TabWindow.SetCloseBitmaps( IDB_CLOSE_16, IDB_CLOSE_ON_16 );

    // Audit report view: position 0 = TAB_INDEX_AUDIT
    m_AuditRichBox.Create( &m_TabWindow );
    m_AuditRichBox.SetFont( VerdanaFont );
    m_AuditRichBox.SetMargins( 10, 10 );
    m_AuditRichBox.SetReadOnly( true );
    m_AuditRichBox.SetTextMaxLength( 0x7FFFFFFE );          // Maximum
    m_AuditRichBox.AddToEventMask( ENM_SELCHANGE | ENM_MOUSEEVENTS );
    m_TabWindow.Add( nullptr, PXS_STRING_EMPTY, PXS_STRING_EMPTY, false, 0, &m_AuditRichBox );

    // Disk details: position 1 = TAB_INDEX_DISKS
    m_DiskRichBox.Create( &m_TabWindow );
    m_DiskRichBox.SetFont( CourierNewFont );
    m_DiskRichBox.SetMargins( 10, 10 );
    m_DiskRichBox.SetReadOnly( true );
    m_DiskRichBox.SetTextMaxLength( 0x7FFFFFFE );          // Maximum
    m_DiskRichBox.AddToEventMask( ENM_SELCHANGE | ENM_MOUSEEVENTS );
    m_TabWindow.Add( nullptr, PXS_STRING_EMPTY, PXS_STRING_EMPTY, true, 0, &m_DiskRichBox );
    m_TabWindow.SetTabVisible( TAB_INDEX_DISKS, false );

    // Display details: position 2 = TAB_INDEX_DISPLAYS
    m_DisplayRichBox.Create( &m_TabWindow );
    m_DisplayRichBox.SetFont( CourierNewFont );
    m_DisplayRichBox.SetMargins( 10, 10 );
    m_DisplayRichBox.SetReadOnly( true );
    m_DisplayRichBox.SetTextMaxLength( 0x7FFFFFFE );        // Maximum
    m_DisplayRichBox.AddToEventMask( ENM_SELCHANGE | ENM_MOUSEEVENTS );
    m_TabWindow.Add( nullptr, PXS_STRING_EMPTY, PXS_STRING_EMPTY, true, 0, &m_DisplayRichBox );
    m_TabWindow.SetTabVisible( TAB_INDEX_DISPLAYS, false );

    // Firmware: position 3 = TAB_INDEX_FIRMWARE
    m_FirmwareRichBox.Create( &m_TabWindow );
    m_FirmwareRichBox.SetFont( CourierNewFont );
    m_FirmwareRichBox.SetMargins( 10, 10 );
    m_FirmwareRichBox.SetReadOnly( true );
    m_FirmwareRichBox.SetTextMaxLength( 0x7FFFFFFE );    // Maximum
    m_FirmwareRichBox.AddToEventMask( ENM_SELCHANGE | ENM_MOUSEEVENTS );
    m_TabWindow.Add( nullptr, PXS_STRING_EMPTY, PXS_STRING_EMPTY, true, 0, &m_FirmwareRichBox );
    m_TabWindow.SetTabVisible( TAB_INDEX_FIRMWARE, false );

    // CPUID: position 4 = TAB_INDEX_PROCESSORS
    m_ProcessorRichBox.Create( &m_TabWindow );
    m_ProcessorRichBox.SetFont( CourierNewFont );
    m_ProcessorRichBox.SetMargins( 10, 10 );
    m_ProcessorRichBox.SetReadOnly( true );
    m_ProcessorRichBox.SetTextMaxLength( 0x7FFFFFFE );          // Maximum
    m_ProcessorRichBox.AddToEventMask( ENM_SELCHANGE | ENM_MOUSEEVENTS );
    m_TabWindow.Add( nullptr, PXS_STRING_EMPTY, PXS_STRING_EMPTY, true, 0, &m_ProcessorRichBox );
    m_TabWindow.SetTabVisible( TAB_INDEX_PROCESSORS, false );

    // Software details: position 5 = TAB_INDEX_SOFTWARE
    m_SoftwareRichBox.Create( &m_TabWindow );
    m_SoftwareRichBox.SetFont( CourierNewFont );
    m_SoftwareRichBox.SetMargins( 10, 10 );
    m_SoftwareRichBox.SetReadOnly( true );
    m_SoftwareRichBox.SetTextMaxLength( 0x7FFFFFFE );       // Maximum
    m_SoftwareRichBox.AddToEventMask( ENM_SELCHANGE | ENM_MOUSEEVENTS );
    m_TabWindow.Add( nullptr, PXS_STRING_EMPTY, PXS_STRING_EMPTY, true, 0, &m_SoftwareRichBox );
    m_TabWindow.SetTabVisible( TAB_INDEX_SOFTWARE, false );

    // Help: position 6 = TAB_INDEX_HELP
    m_HelpRichBox.Create( &m_TabWindow );
    m_HelpRichBox.SetFont( VerdanaFont );
    m_HelpRichBox.SetMargins( 10, 10 );
    m_HelpRichBox.SetReadOnly( true );
    m_HelpRichBox.SetTextMaxLength( 0x7FFFFFFE );  // Maximum
    m_HelpRichBox.AddToEventMask( ENM_SELCHANGE | ENM_MOUSEEVENTS );
    m_TabWindow.Add( nullptr, PXS_STRING_EMPTY, PXS_STRING_EMPTY, true, 0, &m_HelpRichBox );
    m_TabWindow.SetTabVisible( TAB_INDEX_HELP, false );
    if ( g_pApplication )
    {
        g_pApplication->LoadTextDataResource( IDR_WINAUDIT_HELP_RTF, &HelpRichText );
    }
    m_HelpRichBox.SetRichText( HelpRichText );

    // Logger: position 7 = TAB_INDEX_LOG
    m_LogRichBox.Create( &m_TabWindow );
    m_LogRichBox.SetFont( CourierNewFont );
    m_LogRichBox.SetMargins( 10, 10 );
    m_LogRichBox.SetReadOnly( true );
    m_LogRichBox.SetTextMaxLength( 0x7FFFFFFE );            // Maximum
    m_LogRichBox.AddToEventMask( ENM_SELCHANGE | ENM_MOUSEEVENTS );
    m_TabWindow.Add( nullptr, PXS_STRING_EMPTY, PXS_STRING_EMPTY, true, 0, &m_LogRichBox );
    m_TabWindow.SetTabVisible( TAB_INDEX_LOGGER, false );

    // Splitter with categories window of zero width
    m_Splitter.SetStyle( WS_CLIPCHILDREN, true );
    m_Splitter.Create( &m_WorkSpace );
    m_Splitter.SetAppMessageListener( m_hWindow );
    if ( IsRightToLeftReading() )
    {
        m_Splitter.AttachWindows( &m_TabWindow, &m_SidePanel );
        m_Splitter.SetOffsetTwo( 0 );
    }
    else
    {
        m_Splitter.AttachWindows( &m_SidePanel, &m_TabWindow );
        m_Splitter.SetOffset( 0 );
    }

    ///////////////////////////////////////////////////////////////////////////
    // Status bar

    m_StatusBar.SetStyle( WS_CLIPCHILDREN, true );
    m_StatusBar.Create( &m_MainPanel );

    // Panel 0: Web Address
    panelIndex = m_StatusBar.AddPanel( 250 );
    if ( g_pApplication )
    {
        g_pApplication->GetWebSiteURL( &Caption );
    }
    m_StatusBar.SetText( panelIndex, Caption );
    m_StatusBar.SetHyperlink( panelIndex, true );

    // Panel 1: File saved bitmap
    panelIndex = m_StatusBar.AddPanel( 20 );
    m_StatusBar.SetAlignmentX( panelIndex, PXS_CENTER_ALIGNMENT );

    // Panel 2: Logging state bitmap
    panelIndex = m_StatusBar.AddPanel( 20 );
    m_StatusBar.SetAlignmentX( panelIndex, PXS_CENTER_ALIGNMENT );

    // Panel 3: Computer name
    panelIndex = m_StatusBar.AddPanel( 250 );
    SystemInfo.GetComputerNetBiosName( &Caption );
    m_StatusBar.SetText( panelIndex, Caption );

    // Panel 4: EUPL
    Caption = L"European Union Public Licence";
    panelIndex = m_StatusBar.AddPanel( 250 );
    m_StatusBar.SetText( panelIndex, Caption );

    // Progress bar
    m_ProgressBar.Create( &m_StatusBar );

    // Finish up
    SetColours();
    DoLayout();
}

//===============================================================================================//
//  Description:
//      Create a desktop short cut to this executable
//
//  Parameters:
//      None
//
//  Returns:
//     void
//===============================================================================================//
void WinAuditFrame::DesktopShortCut()
{
    Shell   ShellObject;
    String  ShortCutName, ExeDirectory, ExePath, ApplicationName;

    ShortCutName = L"WinAudit.lnk";
    PXSGetExeDirectory( &ExeDirectory );
    PXSGetExePath( &ExePath );
    PXSGetApplicationName( &ApplicationName );
    ShellObject.CreateDesktopShortCut( ShortCutName, ExeDirectory, ExePath, ApplicationName);
}

//===============================================================================================//
//  Description:
//      Create the menu items
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::CreateMenuItems()
{
    String   Caption;
    WindowsInformation WindowsInfo;

    ///////////////////////////////////////////////////////////////////////////
    // Menu Bar

    m_MenuBar.Create( *this );

    ///////////////////////////////////////////////////////////////////////////
    // File Menu

    // The menu bar item and its pop-up
    m_FileMenuBarItem.SetMenuBarItem( true );
    m_FileMenuBarItem.SetParent( m_MenuBar.GetMenuHandle() );
    m_FilePopup.Show( m_MenuBar.GetMenuHandle(), 0 );

    // Configure the menu items
    m_FileAudit.SetCommandID( PXS_APP_MSG_AUDIT_START );
    m_AccelTable.Add( 'N', PXS_APP_MSG_AUDIT_START );

    m_FileSave.SetCommandID( PXS_APP_MSG_SAVE );
    m_AccelTable.Add( 'S', PXS_APP_MSG_SAVE );

    m_FileEmail.SetCommandID( PXS_APP_MSG_SEND_EMAIL );

    m_FileOdbc.SetCommandID( PXS_APP_MSG_ODBC_EXPORT );

    m_FileShortcut.SetCommandID( PXS_APP_MSG_DESKTOP_SHORTCUT );

    m_FileExit.SetCommandID( PXS_APP_MSG_EXIT );

    // Add the items
    m_FilePopup.AddMenuItem( &m_FileAudit );
    m_FilePopup.AddSeparator();
    m_FilePopup.AddMenuItem( &m_FileSave );
    m_FilePopup.AddMenuItem( &m_FileEmail );
    m_FilePopup.AddMenuItem( &m_FileOdbc );
    m_FilePopup.AddSeparator();
    m_FilePopup.AddMenuItem( &m_FileShortcut );
    m_FilePopup.AddSeparator();
    m_FilePopup.AddMenuItem( &m_FileExit );

    ///////////////////////////////////////////////////////////////////////////
    // Edit Menu

    // The menu bar item and its pop-up
    m_EditMenuBarItem.SetMenuBarItem( true );
    m_EditMenuBarItem.SetParent( m_MenuBar.GetMenuHandle() );
    m_EditPopup.Show( m_MenuBar.GetMenuHandle(), 1 );

    // Configure the menu items
    m_EditSelectAll.SetCommandID( PXS_APP_MSG_SELECT_ALL );
    m_AccelTable.Add( 'A', PXS_APP_MSG_SELECT_ALL );

    m_EditCopy.SetCommandID( PXS_APP_MSG_COPY_SELECTION );
    m_AccelTable.Add( 'C', PXS_APP_MSG_COPY_SELECTION );

    m_EditFind.SetCommandID( PXS_APP_MSG_EDIT_FIND );
    m_AccelTable.Add( 'F', PXS_APP_MSG_EDIT_FIND );

    m_EditSetFont.SetCommandID( PXS_APP_MSG_EDIT_SET_FONT );

    m_EditSmallerText.SetCommandID( PXS_APP_MSG_EDIT_SMALLER_TEXT );

    m_EditBiggerText.SetCommandID( PXS_APP_MSG_EDIT_BIGGER_TEXT );

    // Add the items
    m_EditPopup.AddMenuItem( &m_EditSelectAll );
    m_EditPopup.AddMenuItem( &m_EditCopy );
    m_EditPopup.AddMenuItem( &m_EditFind );
    m_EditPopup.AddSeparator();
    m_EditPopup.AddMenuItem( &m_EditSetFont );
    m_EditPopup.AddMenuItem( &m_EditSmallerText );
    m_EditPopup.AddMenuItem( &m_EditBiggerText );

    ///////////////////////////////////////////////////////////////////////////
    // View Menu

    // The menu bar item and its pop-up
    m_ViewMenuBarItem.SetMenuBarItem( true );
    m_ViewMenuBarItem.SetParent( m_MenuBar.GetMenuHandle() );
    m_ViewPopup.Show( m_MenuBar.GetMenuHandle(), 2 );

    // Configure the menu items
    m_ViewToolBar.SetCommandID( PXS_APP_MSG_TOGGLE_TOOLBAR );

    m_ViewStatusBar.SetCommandID( PXS_APP_MSG_TOGGLE_STATUSBAR );

    m_ViewOptions.SetCommandID( PXS_APP_MSG_VIEW_OPTIONS );
    m_AccelTable.Add( 'O', PXS_APP_MSG_VIEW_OPTIONS);

    m_ViewDiskInfo.SetCommandID( PXS_APP_MSG_VIEW_DISK_INFO );

    m_ViewDisplayInfo.SetCommandID( PXS_APP_MSG_VIEW_DISPLAY_INFO );

    m_ViewFirmwareInfo.SetCommandID( PXS_APP_MSG_VIEW_FIRMWARE_INFO );

    m_ViewPolicyInfo.SetCommandID( PXS_APP_MSG_VIEW_POLICY_INFO );

    m_ViewProcessorInfo.SetCommandID( PXS_APP_MSG_VIEW_PROCESSOR_INFO );

    m_ViewSoftwareInfo.SetCommandID( PXS_APP_MSG_VIEW_SOFTWARE_INFO );

    // Add to the menu
    m_ViewPopup.AddMenuItem( &m_ViewToolBar );
    m_ViewPopup.AddMenuItem( &m_ViewStatusBar );
    m_ViewPopup.AddSeparator();
    m_ViewPopup.AddMenuItem( &m_ViewOptions );
    m_ViewPopup.AddSeparator();
    m_ViewPopup.AddMenuItem( &m_ViewDiskInfo );
    m_ViewPopup.AddMenuItem( &m_ViewDisplayInfo );
    m_ViewPopup.AddMenuItem( &m_ViewFirmwareInfo );
    if ( WindowsInfo.GetMajorVersion() >= 6 )   // RSOP requires Vista or newer
    {
        m_ViewPopup.AddMenuItem( &m_ViewPolicyInfo );
    }
    m_ViewPopup.AddMenuItem( &m_ViewProcessorInfo );
    m_ViewPopup.AddMenuItem( &m_ViewSoftwareInfo );

    // Set initial state
    m_ViewToolBar.SetChecked( true );
    m_ViewStatusBar.SetChecked( true );

    ///////////////////////////////////////////////////////////////////////////
    // Language Menus

    // The menu bar item and its pop-up
    m_LanguageMenuBarItem.SetMenuBarItem( true );
    m_LanguageMenuBarItem.SetParent( m_MenuBar.GetMenuHandle() );
    m_LanguagePopup.Show( m_MenuBar.GetMenuHandle(), 3 );

    // Chinese traditional
    m_LanguageChineseTw.SetCommandID( PXS_APP_MSG_LANGUAGE_CHINESE_TW );
    m_LanguageChineseTw.SetBitmap( IDB_FLAG_TW_16 );

    // Czech
    m_LanguageCzech.SetCommandID( PXS_APP_MSG_LANGUAGE_CZECH );
    m_LanguageCzech.SetBitmap( IDB_FLAG_CZ_16 );

    // Danish
    m_LanguageDanish.SetCommandID( PXS_APP_MSG_LANGUAGE_DANISH );
    m_LanguageDanish.SetBitmap( IDB_FLAG_DA_16 );

    // German
    m_LanguageGerman.SetCommandID( PXS_APP_MSG_LANGUAGE_GERMAN );
    m_LanguageGerman.SetBitmap( IDB_FLAG_DE_16 );

    // English
    m_LanguageEnglish.SetCommandID( PXS_APP_MSG_LANGUAGE_ENGLISH );
    m_LanguageEnglish.SetBitmap( IDB_FLAG_GB_16 );

    // Spanish
    m_LanguageSpanish.SetCommandID( PXS_APP_MSG_LANGUAGE_SPANISH );
    m_LanguageSpanish.SetBitmap( IDB_FLAG_ES_16 );

    // Greek
    m_LanguageGreek.SetCommandID( PXS_APP_MSG_LANGUAGE_GREEK );
    m_LanguageGreek.SetBitmap( IDB_FLAG_GR_16 );

    // Fench - Belgian
    m_LanguageBelgian.SetCommandID( PXS_APP_MSG_LANGUAGE_FRENCH_BE );
    m_LanguageBelgian.SetBitmap( IDB_FLAG_BE_16 );

    // French - France
    m_LanguageFrench.SetCommandID( PXS_APP_MSG_LANGUAGE_FRENCH_FR );
    m_LanguageFrench.SetBitmap( IDB_FLAG_FR_16 );

    // Hebrew
    m_LanguageHebrew.SetCommandID( PXS_APP_MSG_LANGUAGE_HEBREW );
    m_LanguageHebrew.SetBitmap( IDB_FLAG_IL_16 );

    // Indonesian
    m_LanguageIndonesian.SetCommandID( PXS_APP_MSG_LANGUAGE_INDONESIAN );
    m_LanguageIndonesian.SetBitmap( IDB_FLAG_ID_16 );

    // Italian
    m_LanguageItalian.SetCommandID( PXS_APP_MSG_LANGUAGE_ITALIAN );
    m_LanguageItalian.SetBitmap( IDB_FLAG_IT_16 );

    // Japanese
    m_LanguageJapanese.SetCommandID( PXS_APP_MSG_LANGUAGE_JAPANESE );
    m_LanguageJapanese.SetBitmap( IDB_FLAG_JP_16 );

    // Korean
    m_LanguageKorean.SetCommandID( PXS_APP_MSG_LANGUAGE_KOREAN );
    m_LanguageKorean.SetBitmap( IDB_FLAG_KR_16 );

    // Hungarian
    m_LanguageHungarian.SetCommandID( PXS_APP_MSG_LANGUAGE_HUNGARIAN );
    m_LanguageHungarian.SetBitmap( IDB_FLAG_HU_16 );

    // Dutch
    m_LanguageDutch.SetCommandID( PXS_APP_MSG_LANGUAGE_DUTCH );
    m_LanguageDutch.SetBitmap( IDB_FLAG_NL_16 );

    // Polish
    m_LanguagePolish.SetCommandID( PXS_APP_MSG_LANGUAGE_POLISH );
    m_LanguagePolish.SetBitmap( IDB_FLAG_PL_16 );

    // Portuguese - Brazil
    m_LanguageBrazilian.SetCommandID( PXS_APP_MSG_LANGUAGE_PORTUGESE_BR );
    m_LanguageBrazilian.SetBitmap( IDB_FLAG_BR_16 );

    // Portuguese
    m_LanguagePortuguese.SetCommandID( PXS_APP_MSG_LANGUAGE_PORTUGESE_PT );
    m_LanguagePortuguese.SetBitmap( IDB_FLAG_PT_16 );

    // Russian
    m_LanguageRussian.SetCommandID( PXS_APP_MSG_LANGUAGE_RUSSIAN );
    m_LanguageRussian.SetBitmap( IDB_FLAG_RU_16 );

    // Slovak
    m_LanguageSlovak.SetCommandID( PXS_APP_MSG_LANGUAGE_SLOVAK );
    m_LanguageSlovak.SetBitmap( IDB_FLAG_SK_16 );

    // Serbian
    m_LanguageSerbian.SetCommandID( PXS_APP_MSG_LANGUAGE_SERBIAN );
    m_LanguageSerbian.SetBitmap( IDB_FLAG_RS_16 );

    // Finnish
    m_LanguageFinnish.SetCommandID( PXS_APP_MSG_LANGUAGE_FINNISH );
    m_LanguageFinnish.SetBitmap( IDB_FLAG_FI_16 );

    // Thai
    m_LanguageThai.SetCommandID( PXS_APP_MSG_LANGUAGE_THAI );
    m_LanguageThai.SetBitmap( IDB_FLAG_TH_16 );

    // Turkish
    m_LanguageTurkish.SetCommandID( PXS_APP_MSG_LANGUAGE_TURKISH );
    m_LanguageTurkish.SetBitmap( IDB_FLAG_TR_16 );

    // Add the the menu
    m_LanguagePopup.AddMenuItem( &m_LanguageChineseTw );
    m_LanguagePopup.AddMenuItem( &m_LanguageCzech );
    m_LanguagePopup.AddMenuItem( &m_LanguageDanish );
    m_LanguagePopup.AddMenuItem( &m_LanguageGerman );
    m_LanguagePopup.AddMenuItem( &m_LanguageEnglish );
    m_LanguagePopup.AddMenuItem( &m_LanguageSpanish );
    m_LanguagePopup.AddMenuItem( &m_LanguageGreek );
    m_LanguagePopup.AddMenuItem( &m_LanguageBelgian );
    m_LanguagePopup.AddMenuItem( &m_LanguageFrench );
    m_LanguagePopup.AddMenuItem( &m_LanguageHebrew );
    m_LanguagePopup.AddMenuItem( &m_LanguageIndonesian );
    m_LanguagePopup.AddMenuItem( &m_LanguageItalian );
    m_LanguagePopup.AddMenuItem( &m_LanguageJapanese );
    m_LanguagePopup.AddMenuItem( &m_LanguageKorean );
    m_LanguagePopup.AddMenuItem( &m_LanguageHungarian );
    m_LanguagePopup.AddMenuItem( &m_LanguageDutch );
    m_LanguagePopup.AddMenuItem( &m_LanguagePolish );
    m_LanguagePopup.AddMenuItem( &m_LanguageBrazilian );
    m_LanguagePopup.AddMenuItem( &m_LanguagePortuguese );
    m_LanguagePopup.AddMenuItem( &m_LanguageRussian );
    m_LanguagePopup.AddMenuItem( &m_LanguageSlovak );
    m_LanguagePopup.AddMenuItem( &m_LanguageSerbian );
    m_LanguagePopup.AddMenuItem( &m_LanguageFinnish );
    m_LanguagePopup.AddMenuItem( &m_LanguageThai );
    m_LanguagePopup.AddMenuItem( &m_LanguageTurkish );

    ///////////////////////////////////////////////////////////////////////////
    // Help Menu

    // The menu bar item and its pop-up
    m_HelpMenuBarItem.SetMenuBarItem( true );
    m_HelpMenuBarItem.SetParent( m_MenuBar.GetMenuHandle() );
    m_HelpPopup.Show( m_MenuBar.GetMenuHandle(), 4 );

    m_HelpUsingWinAudit.SetCommandID( PXS_APP_MSG_HELP );

    m_HelpWebsite.SetCommandID( PXS_APP_MSG_START_BROWSER );
    m_HelpWebsite.SetHyperLink( true );

    m_HelpAbout.SetCommandID( PXS_APP_MSG_ABOUT_DIALOG );
    m_HelpCreateGuid.SetCommandID( PXS_APP_MSG_HELP_CREATE_GUID );
    m_HelpLogFile.SetCommandID( PXS_APP_MSG_TOGGLE_LOGFILE );

    // Add the the menu
    m_HelpPopup.AddMenuItem( &m_HelpUsingWinAudit );
    m_HelpPopup.AddSeparator();
    m_HelpPopup.AddMenuItem( &m_HelpWebsite );
    m_HelpPopup.AddMenuItem( &m_HelpAbout );
    // m_HelpPopup.AddMenuItem( &m_HelpCreateGuid );
    m_HelpPopup.AddSeparator();
    m_HelpPopup.AddMenuItem( &m_HelpLogFile );

    ///////////////////////////////////////////////////////////////////////////
    // Finish up

    // Add the pop-ups to menu bar
    m_MenuBar.AddPopup( m_FileMenuBarItem    , m_FilePopup );
    m_MenuBar.AddPopup( m_EditMenuBarItem    , m_EditPopup );
    m_MenuBar.AddPopup( m_ViewMenuBarItem    , m_ViewPopup );
    m_MenuBar.AddPopup( m_LanguageMenuBarItem, m_LanguagePopup );
    m_MenuBar.AddPopup( m_HelpMenuBarItem    , m_HelpPopup );

    m_AccelTable.Create();
    SetMenuLabels();
    m_MenuBar.Draw();
}

//===============================================================================================//
//  Description:
//      Layout the child windows
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::DoAuditInCommandLineMode()
{
    DWORD  categoryID = 0;
    size_t i = 0, numSwitches, numCategories;
    String CommandLine, FileSwitchValue, ReportSwitchValue, LogSwitchValue;
    String TimestampSwitchValue, LanguageSwitchValue, LogPath, Switch, LocalTimeIso;
    String OutputPath, LogDrive, LogDir, LogFname, LogExt;
    AuditData   Auditor;
    Formatter   Format;
    Directory   DirObject;
    StringArray Switches;
    TArray< DWORD > Categories;
    TArray< AuditRecord > AuditRecords, CategoryRecords;

    // Expand and get the switches, will use the entire string including the exe
    PXSExpandEnvironmentStrings( GetCommandLine(), &CommandLine );
    CommandLine.Trim();
    CommandLine.ToArray( '/', &Switches );
    numSwitches = Switches.GetSize();

    // Scan for help
    for ( i = 0; i < numSwitches; i++ )
    {
        Switch = Switches.Get( i );
        Switch.Trim();
        if ( Switch.StartsWith( L"h", true ) ||
             Switch.StartsWith( L"?", true )  )
        {
            ShowCommandLineHelpMessage();
            return;     // No thing more to do
        }
    }
    PXSGetCommandLineSwitchValues( Switches,
                                   &ReportSwitchValue,
                                   &FileSwitchValue,
                                   &LogSwitchValue, &TimestampSwitchValue, &LanguageSwitchValue );

    // Start logging to a file
    if ( LogSwitchValue.GetLength() )
    {
        // Want full path, send to exe directory if none specified
        if ( LogSwitchValue.IndexOf( PXS_PATH_SEPARATOR, 0 ) == PXS_MINUS_ONE )
        {
            PXSGetExeDirectory( &LogPath );
        }
        LogPath += LogSwitchValue;
        DirObject.SplitPath( LogPath, &LogDrive, &LogDir, &LogFname, &LogExt );
        if ( LogExt.GetLength() == 0 )
        {
            LogPath += L".txt";
        }

        if ( g_pApplication )
        {
            g_pApplication->SetLogLevel( PXS_LOG_LEVEL_VERBOSE );
            g_pApplication->StartLogger( LogPath, false, false );
        }
    }

    // Language file, English only as may want to force this
    if ( g_pApplication && ( LanguageSwitchValue.CompareI( L"en" ) == 0 ) )
    {
        g_pApplication->LoadStringFile( LANG_ENGLISH, 0 );
    }

    // Do the audit
    PXSLogAppInfo1( L"Command Line: '%%1' ", CommandLine );
    LocalTimeIso = Format.LocalTimeInIsoFormat();  // YYYY-MM-DD HH:MM:SS
    SetReportOptions( ReportSwitchValue );
    m_ConfigurationSettings.MakeDataCategoriesArray( &Categories );
    numCategories = Categories.GetSize();
    if ( numCategories == 0 )
    {
        PXSLogAppError( L"No categories for the report were specified. "
                        L"See the /r= switch in the Command Line Usage "
                        L"section of the help. Examples are also given." );
    }
    for ( i = 0; i < numCategories; i++ )
    {
        // Catch errors in individual categories so can continue the job
        try
        {
            // Get the data
            CategoryRecords.RemoveAll();
            categoryID = Categories.Get( i );
            Auditor.GetCategoryRecords( categoryID, LocalTimeIso, &CategoryRecords );
            AuditRecords.Append( CategoryRecords );
        }
        catch ( const Exception& e )
        {
            PXSLogException( e, __FUNCTION__ );
        }
    }

    // Save it
    PXSMakeCommandLineOutputPath( FileSwitchValue,
                                  TimestampSwitchValue, LocalTimeIso, &OutputPath );
    PXSSaveAuditCommandLine( OutputPath, AuditRecords );
}

/*
void WinAuditFrame::DoAuditInCommandLineMode()
{
    DWORD  categoryID = 0;
    size_t i = 0, numSwitches, numCategories;
    String CommandLine, FileSwitchValue, ReportSwitchValue, LogSwitchValue;
    String TimestampSwitchValue, LogPath, Switch, LocalTimeIso, OutputPath;
    String LogDrive, LogDir, LogFname, LogExt;
    AuditData   Auditor;
    Formatter   Format;
    Directory   DirObject;
    StringArray Switches;
    TArray< DWORD > Categories;
    TArray< AuditRecord > AuditRecords, CategoryRecords;

    // Expand and get the switches, will use the entire string including the exe
    PXSExpandEnvironmentStrings( GetCommandLine(), &CommandLine );
    CommandLine.Trim();
    CommandLine.ToArray( '/', &Switches );
    numSwitches = Switches.GetSize();

    // Scan for help
    for ( i = 0; i < numSwitches; i++ )
    {
        Switch = Switches.Get( i );
        Switch.Trim();
        if ( Switch.StartsWith( L"h", true ) ||
             Switch.StartsWith( L"?", true )  )
        {
            ShowCommandLineHelpMessage();
            return;     // No thing more to do
        }
    }
    PXSGetCommandLineSwitchValues( Switches,
                                   &ReportSwitchValue,
                                   &FileSwitchValue, &LogSwitchValue, &TimestampSwitchValue );

    // Start logging to a file
    if ( LogSwitchValue.GetLength() )
    {
        // Want full path, send to exe directory if none specified
        if ( LogSwitchValue.IndexOf( PXS_PATH_SEPARATOR, 0 ) == PXS_MINUS_ONE )
        {
            PXSGetExeDirectory( &LogPath );
        }
        LogPath += LogSwitchValue;
        DirObject.SplitPath( LogPath, &LogDrive, &LogDir, &LogFname, &LogExt );
        if ( LogExt.GetLength() == 0 )
        {
            LogPath += L".txt";
        }

        if ( g_pApplication )
        {
            g_pApplication->SetLogLevel( PXS_LOG_LEVEL_VERBOSE );
            g_pApplication->StartLogger( LogPath, false, false );
        }
    }
    PXSLogAppInfo1( L"Command Line: '%%1' ", CommandLine );

    // Do the audit
    LocalTimeIso = Format.LocalTimeInIsoFormat();  // YYYY-MM-DD HH:MM:SS
    SetReportOptions( ReportSwitchValue );
    m_ConfigurationSettings.MakeDataCategoriesArray( &Categories );
    numCategories = Categories.GetSize();
    if ( numCategories == 0 )
    {
        PXSLogAppError( L"No categories for the report were specified. "
                        L"See the /r= switch in the Command Line Usage "
                        L"section of the help. Examples are also given." );
    }
    for ( i = 0; i < numCategories; i++ )
    {
        // Catch errors in individual categories so can continue the job
        try
        {
            // Get the data
            CategoryRecords.RemoveAll();
            categoryID = Categories.Get( i );
            Auditor.GetCategoryRecords( categoryID, LocalTimeIso, &CategoryRecords );
            AuditRecords.Append( CategoryRecords );
        }
        catch ( const Exception& e )
        {
            PXSLogException( e, __FUNCTION__ );
        }
    }

    // Save it
    PXSMakeCommandLineOutputPath( FileSwitchValue,
                                  TimestampSwitchValue, LocalTimeIso, &OutputPath );
    PXSSaveAuditCommandLine( OutputPath, AuditRecords );
}
*/
//===============================================================================================//
//  Description:
//      Layout the child windows
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::DoLayout()
{
    int  H = 0, toolBarHeight = 0, findTextBarHeight = 0, statusBarHeight = 0;
    SIZE clientSize = { 0, 0 };
    RECT bounds = { 0, 0, 0, 0 };

    if ( m_bCreatedControls == false )
    {
        return;     // Nothing to do
    }
    GetClientSize( &clientSize );

    // The main panel fills the frame's client area
    m_MainPanel.SetBounds( 0, 0, clientSize.cx, clientSize.cy );

    // Tool bar is at the top
    if ( m_ViewToolBar.IsChecked() )
    {
        toolBarHeight = m_ToolBar.GetPreferredHeight();
    }
    m_ToolBar.SetBounds( 0, 0, clientSize.cx, toolBarHeight );
    m_ToolBar.DoLayout();

    // The find text bar is below
    if ( m_FindTextBar.IsVisible() )
    {
        findTextBarHeight = m_FindTextBar.GetPreferredHeight();
        m_FindTextBar.SetBounds( 0, toolBarHeight, clientSize.cx, findTextBarHeight );
    }

    // Status bar is at the bottom
    if ( m_ViewStatusBar.IsChecked() )
    {
        statusBarHeight = m_StatusBar.GetPreferredHeight();
    }
    m_StatusBar.SetBounds( 0, clientSize.cy - statusBarHeight, clientSize.cx, statusBarHeight );
    m_StatusBar.DoLayout();

    // Progress bar spans the status bar
    bounds.left   = 1;
    bounds.top    = 0;
    bounds.right  = clientSize.cx;
    bounds.bottom = statusBarHeight;
    m_ProgressBar.SetBounds( bounds );

    // The workspaces fills in between
    H = clientSize.cy - toolBarHeight - findTextBarHeight - statusBarHeight;
    m_WorkSpace.SetBounds( 0, toolBarHeight + findTextBarHeight, clientSize.cx - 1, H );

    m_FindTextBar.DoLayout();
    m_Splitter.DoLayout();
    m_SidePanel.GetSize( &clientSize );
    m_AuditCategories.SetSize( clientSize );  // Fills the entire side panel
    m_HelpContents.SetSize( clientSize );     // Fills the entire side panel
    m_TabWindow.DoLayout();
    Repaint();
}

//===============================================================================================//
//  Description:
//     Read the ini file
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::ReadIniFile()
{
    File      FileObject;
    DWORD     dword = 0;
    size_t    i = 0, numLines;
    wchar_t*  endptr = nullptr;
    String    Line, IniPath, Name, Value;
    Formatter Format;
    StringArray Lines, Tokens;

    // Path
    PXSGetExeDirectory( &IniPath );
    IniPath += PXS_WINAUDIT_INI;
    if ( FileObject.Exists( IniPath ) == false )
    {
        return;     // Nothing to do
    }
    FileObject.OpenText( IniPath );
    FileObject.ReadLineArray( IniPath, 1000, &Lines );
    FileObject.Close();
    numLines = Lines.GetSize();
    for ( i = 0; i < numLines; i++ )
    {
        Line = Lines.Get( i );
        Line.Trim();
        Line.ToArray( '=', &Tokens );
        if ( 2 == Tokens.GetSize() )
        {
            Name  = Tokens.Get( 0 );
            Value = Tokens.Get( 1 );
            Name.Trim();
            Value.Trim();
            if ( Value.GetLength() )
            {
                if ( Name.CompareI( L"autoStartAudit" ) == 0 )
                {
                    if ( Value.CompareI( L"1" ) == 0 )
                    {
                        m_ConfigurationSettings.autoStartAudit = true;
                    }
                    else
                    {
                        m_ConfigurationSettings.autoStartAudit = false;
                    }
                }
                else if ( Name.CompareI( L"r" ) == 0 )
                {
                    SetReportOptions( Value );
                }
                else if ( Name.CompareI( L"reportShowSQL" ) == 0 )
                {
                    if ( Value.CompareI( L"1" ) == 0 )
                    {
                        m_ConfigurationSettings.reportShowSQL = true;
                    }
                    else
                    {
                        m_ConfigurationSettings.reportShowSQL = false;
                    }
                }
                else if ( Name.CompareI( L"maxErrorRate" ) == 0 )
                {
                    dword = wcstoul( Value.c_str(), &endptr, 10 );
                    if ( dword <= PXS_DB_MAX_ERROR_RATE_MAX )
                    {
                        m_ConfigurationSettings.maxErrorRate = dword;
                    }
                }
                else if ( Name.CompareI( L"maxAffectedRows" ) == 0 )
                {
                    dword = wcstoul( Value.c_str(), &endptr, 10 );
                    if ( ( dword >= PXS_DB_MAX_AFFECTED_ROWS_MIN ) &&
                         ( dword <= PXS_DB_MAX_AFFECTED_ROWS_MAX ) )
                    {
                        m_ConfigurationSettings.maxAffectedRows = dword;
                    }
                }
                else if ( Name.CompareI( L"connectTimeoutSecs" ) == 0 )
                {
                    dword = wcstoul( Value.c_str(), &endptr, 10 );
                    if ( ( dword >= PXS_DB_CONNECT_TIMEOUT_SECS_MIN ) &&
                         ( dword <= PXS_DB_CONNECT_TIMEOUT_SECS_MAX ) )
                    {
                        m_ConfigurationSettings.connectTimeoutSecs = dword;
                    }
                }
                else if ( Name.CompareI( L"queryTimeoutSecs" ) == 0 )
                {
                    dword = wcstoul( Value.c_str(), &endptr, 10 );
                    if ( ( dword >= PXS_DB_QUERY_TIMEOUT_SECS_MIN ) &&
                         ( dword <= PXS_DB_QUERY_TIMEOUT_SECS_MAX ) )
                    {
                        m_ConfigurationSettings.queryTimeoutSecs = dword;
                    }
                }
                else if ( Name.CompareI( L"reportMaxRecords" ) == 0 )
                {
                    dword = wcstoul( Value.c_str(), &endptr, 10 );
                    if ( ( dword >= PXS_REPORT_MAX_RECORDS_MIN ) &&
                         ( dword <= PXS_REPORT_MAX_RECORDS_MAX ) )
                    {
                        m_ConfigurationSettings.reportMaxRecords = dword;
                    }
                }
                else if ( Name.CompareI( L"DBMS" ) == 0 )
                {
                    m_ConfigurationSettings.DBMS = Value;
                }
                else if ( Name.CompareI( L"DatabaseName" ) == 0 )
                {
                    m_ConfigurationSettings.DatabaseName = Value;
                }
                else if ( Name.CompareI( L"MySqlDriver" ) == 0 )
                {
                    m_ConfigurationSettings.MySqlDriver = Value;
                }
                else if ( Name.CompareI( L"PostgreSqlDriver" ) == 0 )
                {
                    m_ConfigurationSettings.PostgreSqlDriver = Value;
                }
                else if ( Name.CompareI( L"LastReportName" ) == 0 )
                {
                    m_ConfigurationSettings.LastReportName = Value;
                }
                else if ( Name.CompareI( L"AppFrameBounds" ) == 0 )
                {
                    SetFrameBounds( Value );
                }
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Set the state of the controls depending
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::SetControlsAndMenus()
{
    String Label, LogFilePath;
    WindowsInformation WindowsInfo;

    if ( m_bCreatedControls == false )
    {
        throw FunctionException( L"m_bCreatedControls", __FUNCTION__ );
    }

    // Auditing
    m_FileAudit.SetEnabled( !m_bAuditing );
    m_EditMenuBarItem.SetEnabled( !m_bAuditing );
    m_ViewMenuBarItem.SetEnabled( !m_bAuditing );
    m_LanguageMenuBarItem.SetEnabled( !m_bAuditing );
    m_MenuBar.Draw();
    m_AuditButton.SetEnabled( !m_bAuditing );
    m_StopButton.SetEnabled( m_bAuditing );
    m_ProgressBar.SetVisible( m_bAuditing );

    // Edit menu
    SetEditMenu();

    // Report
    m_FileSave.SetEnabled( !m_bAuditing );
    m_FileEmail.SetEnabled( !m_bAuditing );
    m_FileOdbc.SetEnabled( !m_bAuditing );
    m_SaveButton.SetEnabled( !m_bAuditing );
    m_EmailButton.SetEnabled( !m_bAuditing );
    if ( m_bCreatedReport )
    {
        if ( m_bSavedReport )
        {
            m_StatusBar.SetBitmap( 1, IDB_SAVE_16 );
        }
        else
        {
            m_StatusBar.SetBitmap( 1, IDB_NOT_SAVED_16 );
        }
    }
    else
    {
        m_StatusBar.SetBitmap( 1, 0 );
    }

    // Logging
    if ( m_HelpLogFile.IsChecked() )
    {
        PXSGetResourceString( PXS_IDS_1054_STOP_LOGGING, &Label );
        m_HelpLogFile.SetLabel( Label );
        m_StatusBar.SetBitmap( 2, IDB_LOG_FILE_16 );
    }
    else
    {
        PXSGetResourceString( PXS_IDS_1053_START_LOGGING, &Label );
        m_HelpLogFile.SetLabel( Label );
        m_StatusBar.SetBitmap( 2, 0 );
    }

    // Pre-Install Environment
    if ( WindowsInfo.IsPreInstall() )
    {
        m_FileShortcut.SetEnabled( false );
    }
}

//===============================================================================================//
//  Description:
//     Set the user interface language
//
//  Parameters:
//      primaryLanguage - the primary language id
//      subLanguage     - the sublanguage id
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::SetUserInterfaceLanguage( WORD primaryLanguage, WORD subLanguage )
{
    bool  rtlReading = false;

    if ( g_pApplication == nullptr )
    {
        throw NullException( L"g_pApplication", __FUNCTION__ );
    }
    g_pApplication->LoadStringFile( primaryLanguage, subLanguage );

    // Right to left
    if ( primaryLanguage == LANG_HEBREW  )
    {
        rtlReading = true;
    }
    SetRightToLeftReading( rtlReading );
    m_MainPanel.SetRightToLeftReading( rtlReading );
    m_ToolBar.SetRightToLeftReading( rtlReading );
    m_FindTextBar.SetRightToLeftReading( rtlReading );
    m_Splitter.SetRightToLeftReading( rtlReading );
    m_SidePanel.SetRightToLeftReading( rtlReading );
    m_TabWindow.SetRightToLeftReading( rtlReading );
    m_AuditCategories.SetRightToLeftReading( rtlReading );
    m_HelpContents.SetRightToLeftReading( rtlReading );
    m_AuditRichBox.SetRightToLeftReading( rtlReading );
    m_DiskRichBox.SetRightToLeftReading( rtlReading );
    m_DisplayRichBox.SetRightToLeftReading( rtlReading );
    m_FirmwareRichBox.SetRightToLeftReading( rtlReading );
    m_ProcessorRichBox.SetRightToLeftReading( rtlReading );
    m_SoftwareRichBox.SetRightToLeftReading( rtlReading );
    m_HelpRichBox.SetRightToLeftReading( rtlReading );
    m_AboutRichBox.SetRightToLeftReading( rtlReading );
    m_LogRichBox.SetRightToLeftReading( rtlReading );
    m_StatusBar.SetRightToLeftReading( rtlReading );
    m_ProgressBar.SetRightToLeftReading( rtlReading );

    // Set the splitter's panels
    if ( rtlReading )
    {
        m_Splitter.AttachWindows( &m_TabWindow, &m_SidePanel );
        m_Splitter.SetOffsetTwo( 0 );
    }
    else
    {
        m_Splitter.AttachWindows( &m_SidePanel, &m_TabWindow );
        m_Splitter.SetOffset( 0 );
    }

    // Recreate
    DestroyMenuItems();
    SetMenuLabels();
    SetControlLabels();
    CreateMenuItems();
    SetColours();
    DoLayout();
    Repaint();
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Handle application specific events, i.e. above WM_APP.
//
//  Parameters:
//      uMsg   - the message number, should be greater than WM_APP
//      wParam - message specific data
//      lParam - the window sending the message
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::AppMessageEvent( UINT uMsg, WPARAM /* wParam */, LPARAM lParam )
{
    try
    {
        switch( uMsg )
        {
            default:
                break;

            case PXS_APP_MSG_AUDIT_THREAD_DONE:
                AuditThreadDone();
                break;

            case PXS_APP_MSG_AUDIT_THREAD_UPDATE:
                AuditThreadUpdate();
                break;

            case PXS_APP_MSG_LOGGER_UPDATE:
                UpdateLogWindow();
                break;

            case PXS_APP_MSG_HIDE_WINDOW:
                HandleHideWindow( (HWND)lParam );
                break;
        }
    }
    catch ( const Exception& e )
    {
        PXSShowExceptionDialog( e, m_hWindow );
    }
}

//===============================================================================================//
//  Description:
//      Handle WM_CLOSE event.
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::CloseEvent()
{
    WaitCursor Wait;

    // Make sure redraw where the menu was
    try
    {
        RedrawAllNow();
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }

    // Global stop, if anyone is polling this flag
    try
    {
        if ( g_pApplication )
        {
            g_pApplication->SetStopBackgroundTasks( true );
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }

    // Save the settings, continue on error
    try
    {
        WriteIniFile();
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }

    // Stop any file I/O
    try
    {
        m_AuditThread.Stop();
        m_AuditThread.Join();
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }

    // Destroy the menus, best to do this before destroying the frame otherwise
    // its associated menu handles will be invalid as they are automatically
    // destroyed by the system
    try
    {
        DestroyMenuItems();
    }
    catch ( const Exception& e)
    {
        PXSLogException( e, __FUNCTION__ );
    }
    DestroyWindowHandle();
}

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
LRESULT WinAuditFrame::CommandEvent( WPARAM wParam, LPARAM lParam )
{
    bool    visible  = false;
    WORD    id       = LOWORD(wParam);  // Item, control, or accelerator
    WORD    msg      = HIWORD(wParam);
    HWND    hWnd     = reinterpret_cast<HWND>( lParam );    // Window handle
    size_t  tabIndex = 0;
    LRESULT result   = 0;

    // Ensure controls have been created
    if ( m_bCreatedControls == false )
    {
        return result;
    }

    switch( id )
    {
        default:

            // No id, match on window handle
            if ( ( msg  == EN_CHANGE ) &&
                 ( hWnd == m_FindTextBar.GetTextFieldHwnd() ) )
            {
                // Search from current position
                FindText( true, true );
            }
            else
            {
                result = 1;    // Not handled
            }
            break;

        case PXS_APP_MSG_AUDIT_START:
            AuditStart();
            break;

        case PXS_APP_MSG_SAVE:
            Save();
            break;

        case PXS_APP_MSG_SEND_EMAIL:
            SendEmail();
            break;

        case PXS_APP_MSG_ODBC_EXPORT:
            ShowOdbcExport();
            break;

        case PXS_APP_MSG_DESKTOP_SHORTCUT:
            DesktopShortCut();
            break;

        case PXS_APP_MSG_EXIT:
            CloseEvent();
            break;

        case PXS_APP_MSG_SELECT_ALL:
            SelectAllContents();
            break;

        case PXS_APP_MSG_COPY_SELECTION:
            CopySelection();
            break;

        case PXS_APP_MSG_EDIT_FIND:
            visible = !m_FindTextBar.IsVisible();
            m_FindTextBar.SetVisible( visible );
            m_EditFind.SetChecked( visible );  // Sync with SetVisible
            if ( visible )
            {
                m_FindTextBar.FocusOnTextBox();
            }
            DoLayout();

            break;

        case PXS_APP_MSG_EDIT_SET_FONT:
            SelectFont();
            break;

        case PXS_APP_MSG_EDIT_BIGGER_TEXT:
            SetBiggerSmallerText( true );
            break;

        case PXS_APP_MSG_EDIT_SMALLER_TEXT:
            SetBiggerSmallerText( false );
            break;

        case PXS_APP_MSG_TOGGLE_TOOLBAR:
            m_ViewToolBar.SetChecked( !m_ViewToolBar.IsChecked() );
            DoLayout();
            break;

        case PXS_APP_MSG_TOGGLE_STATUSBAR:
            m_ViewStatusBar.SetChecked( !m_ViewStatusBar.IsChecked() );
            DoLayout();
            break;

        case PXS_APP_MSG_VIEW_OPTIONS:
            ViewOptionsDialog();
            break;

        case PXS_APP_MSG_VIEW_DISK_INFO:
            ViewDiskInformation();
            break;

        case PXS_APP_MSG_VIEW_DISPLAY_INFO:
            ViewDisplayInformation();
            break;

        case PXS_APP_MSG_VIEW_FIRMWARE_INFO:
            ViewFirmwareInformation();
            break;

        case PXS_APP_MSG_VIEW_POLICY_INFO:
            ViewPolicyInformation();
            break;

        case PXS_APP_MSG_VIEW_PROCESSOR_INFO:
            ViewProcessorInformation();
            break;

        case PXS_APP_MSG_VIEW_SOFTWARE_INFO:
            ViewSoftwareInformation();
            break;

        case PXS_APP_MSG_LANGUAGE_CHINESE_TW:
            SetUserInterfaceLanguage( LANG_CHINESE, SUBLANG_CHINESE_TRADITIONAL );
            break;

        case PXS_APP_MSG_LANGUAGE_CZECH:
            SetUserInterfaceLanguage( LANG_CZECH, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_DANISH:
            SetUserInterfaceLanguage( LANG_DANISH, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_GERMAN:
            SetUserInterfaceLanguage( LANG_GERMAN, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_ENGLISH:
            SetUserInterfaceLanguage( LANG_ENGLISH, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_SPANISH:
            SetUserInterfaceLanguage( LANG_SPANISH, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_GREEK:
            SetUserInterfaceLanguage( LANG_GREEK, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_FRENCH_BE:
            SetUserInterfaceLanguage( LANG_FRENCH, SUBLANG_FRENCH_BELGIAN );
            break;

        case PXS_APP_MSG_LANGUAGE_FRENCH_FR:
            SetUserInterfaceLanguage( LANG_FRENCH, SUBLANG_FRENCH );
            break;

        case PXS_APP_MSG_LANGUAGE_HEBREW:
            SetUserInterfaceLanguage( LANG_HEBREW, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_INDONESIAN:
            SetUserInterfaceLanguage( LANG_INDONESIAN, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_ITALIAN:
            SetUserInterfaceLanguage( LANG_ITALIAN, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_JAPANESE:
            SetUserInterfaceLanguage( LANG_JAPANESE, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_KOREAN:
            SetUserInterfaceLanguage( LANG_KOREAN, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_HUNGARIAN:
            SetUserInterfaceLanguage( LANG_HUNGARIAN, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_DUTCH:
            SetUserInterfaceLanguage( LANG_DUTCH, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_POLISH:
            SetUserInterfaceLanguage( LANG_POLISH, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_PORTUGESE_BR:
            SetUserInterfaceLanguage( LANG_PORTUGUESE, SUBLANG_PORTUGUESE_BRAZILIAN );
            break;

        case PXS_APP_MSG_LANGUAGE_PORTUGESE_PT:
            SetUserInterfaceLanguage( LANG_PORTUGUESE, SUBLANG_PORTUGUESE );
            break;

        case PXS_APP_MSG_LANGUAGE_RUSSIAN:
            SetUserInterfaceLanguage( LANG_RUSSIAN, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_SLOVAK:
            SetUserInterfaceLanguage( LANG_SLOVAK, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_SERBIAN:
            SetUserInterfaceLanguage( LANG_SERBIAN, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_FINNISH:
            SetUserInterfaceLanguage( LANG_FINNISH, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_THAI:
            SetUserInterfaceLanguage( LANG_THAI, 0 );
            break;

        case PXS_APP_MSG_LANGUAGE_TURKISH:
            SetUserInterfaceLanguage( LANG_TURKISH, 0 );
            break;

        case PXS_APP_MSG_HELP:
            m_HelpUsingWinAudit.SetChecked( !m_HelpUsingWinAudit.IsChecked() );
            ShowHideHelp( m_HelpUsingWinAudit.IsChecked() );
            break;

        case PXS_APP_MSG_START_BROWSER:
            NavigateToWebSite();
            break;

        case PXS_APP_MSG_ABOUT_DIALOG:
            ShowAboutDialog();
            break;

        case PXS_APP_MSG_HELP_CREATE_GUID:
            ShowCreateGuid();
            break;

        case PXS_APP_MSG_TOGGLE_LOGFILE:
            m_HelpLogFile.SetChecked( !m_HelpLogFile.IsChecked() );
            StartStopLogger( m_HelpLogFile.IsChecked() );
            break;

        case PXS_APP_MSG_AUDIT_COMMAND_LINE:
            DoAuditInCommandLineMode();
            break;

        case PXS_APP_MSG_SPLITTER:
            DoLayout();
            break;

        case PXS_APP_MSG_ITEM_SELECTED:
            HandleItemSelected( hWnd );
            break;

        case PXS_APP_MSG_BUTTON_CLICK:
            HandleButtonClick( hWnd );
            break;

        case PXS_APP_MSG_REMOVE_ITEM:

            if ( hWnd == m_TabWindow.GetHwnd() )
            {
                tabIndex = m_TabWindow.GetSelectedTabIndex();
                m_TabWindow.SetTabVisible( tabIndex, false );
                if ( tabIndex == TAB_INDEX_HELP )
                {
                    m_AuditCategories.SetVisible( true );
                    m_HelpContents.SetVisible( false );
                    m_HelpUsingWinAudit.SetChecked( false );
                }
                else if ( tabIndex == TAB_INDEX_LOGGER )
                {
                    StartStopLogger( false );
                }
            }
            break;
    }

    return result;
}

//===============================================================================================//
//  Description:
//      Handle WM_COPYDATA event.
//
//  Parameters:
//      hWnd - handle to window sending the data
//      pCds - pointer to a COPYDATASTRUCT holding the data
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::CopyDataEvent( HWND hWnd, const COPYDATASTRUCT* pCds )
{
    String Bookmark;

    if ( pCds == nullptr )
    {
        return;
    }
    Bookmark = reinterpret_cast<LPCWSTR>( pCds->lpData );

    if ( hWnd == m_AuditCategories.GetHwnd() )
    {
        m_TabWindow.SetSelectedTabIndex( TAB_INDEX_AUDIT );
        if ( m_AuditRichBox.FindText( Bookmark, true, true, true ) == false )
        {
            PXSLogAppWarn1( L"Audit bookmark '%%1' not found.", Bookmark );
        }
        SetFocus( hWnd );
    }
    else if ( hWnd == m_HelpContents.GetHwnd() )
    {
        m_TabWindow.SetSelectedTabIndex( TAB_INDEX_HELP );
        if ( m_HelpRichBox.FindText( Bookmark, true, true, true ) == false )
        {
            PXSLogAppWarn1( L"Help bookmark '%%1' not found.", Bookmark );
        }
        SetFocus( hWnd );
    }
}

//===============================================================================================//
//  Description:
//      Handle the WM_HELP message
//
//  Parameters:
//      None
//
//  Remarks:
//      Glitch! As showing the help in a tab rather than in a separate frame,
//      if a help event is raised from a dialog (e.g. F1 key) the user will not
//      be able to access the tab.
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::HelpEvent( const HELPINFO* /* pHelpInfo */ )
{
    ShowHideHelp( true );
}

//===============================================================================================//
//  Description:
//      Handle WM_NOTIFY event.
//
//  Parameters:
//      pNmhdr - pointer to a NMHDR structure
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::NotifyEvent( const NMHDR* pNmhdr )
{
    const MSGFILTER* pMsgFilter = nullptr;

    if ( pNmhdr == nullptr )
    {
        return;     // Nothing to do
    }

    if ( pNmhdr->code == EN_SELCHANGE )
    {
        // Rich edit selection change
        SetEditMenu();
    }
    else if ( pNmhdr->code == EN_MSGFILTER )
    {
        // Show the pop-up menu,  pNmhdr is first member of MSGFILTER
        pMsgFilter = reinterpret_cast<const MSGFILTER*>( pNmhdr );
        if ( pMsgFilter->msg == WM_RBUTTONUP )
        {
            POINT point = { 0, 0 };
            HMENU hEditMenu = m_EditPopup.GetMenuHandle();
            if ( hEditMenu )
            {
                GetCursorPos( &point );
                TrackPopupMenu( hEditMenu, 0, point.x, point.y, 0, m_hWindow, nullptr );
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Handle WM_SETTINGCHANGE event.
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::SettingChangeEvent()
{
    // Recreate hi colour parts of the application
    try
    {
        DestroyMenuItems();
        SetMenuLabels();
        SetControlLabels();
        CreateMenuItems();
        SetColours();
        DoLayout();
        Repaint();
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Create the menu items
//
//  Parameters:
//      None
//
//  Remarks:
//      Need to recreate all menus and items because the system
//      only evaluates the size of pop-up menus when they are first
//      shown.
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::DestroyMenuItems()
{
    // Seems reasonable to dissociate the pop-ups from the menu bar before
    // destroying them.
    m_MenuBar.RemoveAllPopupMenus();
    m_FilePopup.Destroy();
    m_EditPopup.Destroy();
    m_ViewPopup.Destroy();
    m_LanguagePopup.Destroy();
    m_HelpPopup.Destroy();

    m_MenuBar.Destroy();
    m_AccelTable.Destroy();
}

//===============================================================================================//
//  Description:
//     Find contents of the help tree view
//
//  Parameters:
//      None
//
//  Remarks:
//       Keep in sync with the help documentation
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::FillHelpContents()
{
    size_t i = 0;
    String Caption, StringData;
    TreeViewItem HelpItem;

    struct _HELP_CONTENTS
    {
        bool    isNode;
        BYTE    indent;
        LPCWSTR pszTitle;
    } HelpContents[] =
        { { false, 0, L"1) Introduction"                   },
          { false, 0, L"2) What's New"                     },
          { true , 0, L"3) Install/Remove WinAudit"        },
          { false, 1, L"3.1) System Requirements"          },
          { false, 1, L"3.2) Installing WinAudit"          },
          { false, 1, L"3.3) Removing WinAudit"            },
          { false, 0, L"4) Frequently Asked Questions"     },
          { false, 0, L"5) Auditing Your Computer"         },
          { true , 0, L"6) Audit Report"                   },
          { false, 1, L"6.1) System Overview"              },
          { false, 1, L"6.2) Installed Software"           },
          { false, 1, L"6.3) Operating System"             },
          { false, 1, L"6.4) Peripherals"                  },
          { false, 1, L"6.5) Security"                     },
          { true , 1, L"6.6) Groups and Users"             },
          { false, 2, L"6.6.1) Groups"                     },
          { false, 2, L"6.6.2) Group Members"              },
          { false, 2, L"6.6.3) Group Policy"               },
          { false, 2, L"6.6.4) Users"                      },
          { false, 2, L"6.7) Scheduled Tasks"              },
          { false, 2, L"6.8) Uptime Statistics"            },
          { false, 2, L"6.9) Error Logs"                   },
          { false, 2, L"6.10) Environment Variables"       },
          { false, 2, L"6.11) Regional Settings"           },
          { true , 2, L"6.12) Windows(R) Network"          },
          { false, 3, L"6.12.1) Network Files"             },
          { false, 3, L"6.12.2) Network Sessions"          },
          { false, 3, L"6.12.3) Network Shares"            },
          { true , 1, L"6.13) Network TCP/IP"              },
          { false, 2, L"6.13.1) Network Adapters"          },
          { false, 2, L"6.13.2) Open Ports"                },
          { false, 1, L"6.14) Hardware Devices"            },
          { false, 1, L"6.15) Display Capabilities"        },
          { false, 1, L"6.16) Display Adapters"            },
          { false, 1, L"6.17) Installed Printers"          },
          { false, 1, L"6.18) BIOS Version"                },
          { true , 1, L"6.19) System Management"           },
          { false, 2, L"6.19.1) BIOS Details"              },
          { false, 2, L"6.19.2) System Information"        },
          { false, 2, L"6.19.3) Base Board"                },
          { false, 2, L"6.19.4) Chassis Information"       },
          { false, 2, L"6.19.5) Processor"                 },
          { false, 2, L"6.19.6) Memory Controller"         },
          { false, 2, L"6.19.7) Memory Module"             },
          { false, 2, L"6.19.8) Cache"                     },
          { false, 2, L"6.19.9) Port Connector"            },
          { false, 2, L"6.19.10) System Slot"              },
          { false, 2, L"6.19.11) Memory Array Information" },
          { false, 2, L"6.19.12) Memory Device"            },
          { false, 1, L"6.20) Processors"                  },
          { false, 1, L"6.21) Memory"                      },
          { false, 1, L"6.22) Physical Disks"              },
          { false, 1, L"6.23) Drives"                      },
          { false, 1, L"6.24) Communication Ports"         },
          { false, 1, L"6.25) Startup Programs"            },
          { false, 1, L"6.26) Services"                    },
          { false, 1, L"6.27) Running Programs"            },
          { true , 1, L"6.28) ODBC Information"            },
          { false, 2, L"6.28.1) ODBC Data Sources"         },
          { false, 2, L"6.28.2) ODBC Drivers"              },
          { false, 1, L"6.29) OLE DB Drivers"              },
          { false, 1, L"6.30) Software Metering"           },
          { false, 1, L"6.31) User Logon Statistics"       },
          { false, 0, L"7) Saving the Audit"               },
          { false, 0, L"8) Sending by E-Mail"              },
          { true , 0, L"9) Export to Database"             },
          { true , 1, L"9.1) Administration - DBA"         },
          { false, 2, L"9.1.1) Create the Database"        },
          { false, 2, L"9.1.2) Client Connection"          },
          { false, 2, L"9.1.3) Data Maintenance"           },
          { false, 2, L"9.1.4) Reporting"                  },
          { true , 0, L"10) Command Line Usage"            },
          { false, 1, L"10.1) Command line examples"       },
          { false, 0, L"11) Privacy and Security"          },
          { false, 0, L"12) Contact Information"           },
          { false, 0, L"13) European Union Public Licence" },
          { false, 0, L"14) Acknowledgements"              } };

    for ( i = 0; i < ARRAYSIZE( HelpContents ); i++ )
    {
        Caption    = HelpContents[ i ].pszTitle;
        StringData = HelpContents[ i ].pszTitle;

        HelpItem.Reset();
        HelpItem.SetNode( HelpContents[ i ].isNode );
        HelpItem.SetIndent( HelpContents[ i ].indent );
        HelpItem.SetLabel( Caption );
        HelpItem.SetStringData( StringData );
        m_HelpContents.AddItem( HelpItem );
    }
}

//===============================================================================================//
//  Description:
//     Find the text shown on the Find Text Bar in the currently selected tab
//
//  Parameters:
//      forward - true to find forward, else backward
//      fromSelectionStart - search begins at start of current selection
//                           otherwise from end of selection
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::FindText( bool forward, bool fromSelectionStart )
{
    bool    caseSensitive = false;
    String  Text, ClassName;
    Window* pWindow;
    RichEditBox* pRichEditBox;

    m_FindTextBar.ShowTextNotFoundLabel( false );
    m_FindTextBar.GetSearchParameters( &Text, &caseSensitive );
    if ( Text.IsEmpty() )
    {
        return;     // Nothing to to
    }

    pWindow = m_TabWindow.GetSelectedTabWindow();
    if ( pWindow == nullptr )
    {
       return;  // No tab selected so nothing to do
    }

    // Verify the window is a rich edit
    pWindow->GetWndClassExClassName( &ClassName );
    if (  ClassName.CompareI( MSFTEDIT_CLASS ) )
    {
        throw SystemException( ERROR_INVALID_WINDOW_STYLE, L"MSFTEDIT_CLASS", __FUNCTION__ );
    }
    pRichEditBox = static_cast< RichEditBox* >( pWindow );
    if ( false == pRichEditBox->FindText( Text, caseSensitive, forward, fromSelectionStart ) )
    {
        m_FindTextBar.ShowTextNotFoundLabel( true );
    }
}

//===============================================================================================//
//  Description:
//      Get if want the audit to run automatically on program start up
//
//  Parameters:
//      none
//
//  Returns:
//      true for auto start, otherwise false
//===============================================================================================//
bool WinAuditFrame::GetAutoStartAudit()
{
    return m_ConfigurationSettings.autoStartAudit;
}

//===============================================================================================//
//  Description:
//     Handle a button click message
//
//  Parameters:
//      hWnd - window that raised the message
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::HandleButtonClick( HWND hWnd )
{
    if ( hWnd == m_FindTextBar.GetFindFowardHwnd() )
    {
        FindText( true, false );
    }
    else if ( hWnd == m_FindTextBar.GetFindBackwardHwnd() )
    {
        FindText( false, false );
    }
    else if ( hWnd == m_AuditButton.GetHwnd() )
    {
        AuditStart();
    }
    else if ( hWnd == m_StopButton.GetHwnd() )
    {
        if ( g_pApplication )
        {
            g_pApplication->SetStopBackgroundTasks( true );
        }
        m_bRunAuditThreadMT = FALSE;
    }
    else if ( hWnd == m_OptionsButton.GetHwnd() )
    {
        ViewOptionsDialog();
    }
    else if ( hWnd == m_SaveButton.GetHwnd() )
    {
        Save();
    }
    else if ( hWnd == m_EmailButton.GetHwnd() )
    {
        SendEmail();
    }
    else if ( hWnd == m_HelpButton.GetHwnd() )
    {
        m_HelpUsingWinAudit.SetChecked( true );
        ShowHideHelp( true );
    }
}

//===============================================================================================//
//  Description:
//     Handle the hide window PXS_APP_MSG_HIDE_WINDOW message
//
//  Parameters:
//      hWnd - window that raised the message
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::HandleHideWindow( HWND hWnd )
{
    if ( hWnd == m_ToolBar.GetHwnd() )
    {
        m_ViewToolBar.SetChecked( false );
        DoLayout();
    }
    else if ( ( hWnd == m_AuditCategories.GetHwnd() ) ||
              ( hWnd == m_HelpContents.GetHwnd()    )  )
    {
        if ( IsRightToLeftReading() )
        {
            m_Splitter.SetOffsetTwo( 0 );
        }
        else
        {
            m_Splitter.SetOffset( 0 );
        }
        DoLayout();
    }
    else if ( hWnd == m_FindTextBar.GetHwnd() )
    {
        m_FindTextBar.SetVisible( false );
        m_EditFind.SetChecked( false );  // Sync with SetVisible
        DoLayout();
    }
}

//===============================================================================================//
//  Description:
//     Handle the item selected message
//
//  Parameters:
//      hWnd - window that raised the message
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::HandleItemSelected( HWND hWnd )
{
    // Process tab selected
    if ( hWnd == m_TabWindow.GetHwnd() )
    {
        // Show/Hide the audit or help contents tree
        if ( TAB_INDEX_AUDIT == m_TabWindow.GetSelectedTabIndex() )
        {
            m_AuditCategories.SetVisible( true );
            m_HelpContents.SetVisible( false );
        }
        else if ( TAB_INDEX_HELP == m_TabWindow.GetSelectedTabIndex() )
        {
            m_AuditCategories.SetVisible( false );
            m_HelpContents.SetVisible( true );
        }
        SetEditMenu();   // Set the Edit menu for rich edit controls
    }
}

//===============================================================================================//
//  Description:
//     Start the browser and go the application's website
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::NavigateToWebSite()
{
    Shell  ShellObject;
    String WebSiteURL;

    if ( g_pApplication == nullptr )
    {
        throw NullException( L"g_pApplication", __FUNCTION__ );
    }
    g_pApplication->GetWebSiteURL( &WebSiteURL );
    ShellObject.NavigateToUrl( WebSiteURL );
}

//===============================================================================================//
//  Description:
//     Prompt the user to save the contents of the currently selected tab
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::Save()
{
    DWORD     filterIndex = 0;
    size_t    tabIndex = m_TabWindow.GetSelectedTabIndex();
    String    ErrorMessage, ComputerName, FilePath;
    Formatter Format;
    SystemInformation SystemInfo;
    LPCWSTR pszFileName   = L"";
    LPCWSTR pszFilter     = L"All Files\0*.*\0\0";
    LPCWSTR pszExtensions = nullptr;

    if ( tabIndex == TAB_INDEX_AUDIT )
    {
        SystemInfo.GetComputerNetBiosName( &ComputerName );
        pszFilter   = L"Comma Delimited (*.csv)\0*.csv\0"
                      L"Comma Delimited 2 (*.csv2)\0*.csv2\0"
                      L"Rich Text (*.rtf)\0*.rtf\0"
                      L"Web Page (*.htm;*.html)\0*.htm;*.html\0\0";
        pszExtensions = L"csv csv2 rtf html";
        if ( m_uSaveAuditFilterIndex )
        {
            filterIndex = m_uSaveAuditFilterIndex;
        }
        else
        {
            filterIndex = DATA_FORMAT_DEFAULT;
        }

        if ( filterIndex == DATA_FORMAT_CSV )
        {
            ComputerName += L".csv";
        }
        else if ( filterIndex == DATA_FORMAT_CSV2 )
        {
            ComputerName += L".csv2";
        }
        else if ( filterIndex == DATA_FORMAT_RTF )
        {
            ComputerName += L".rtf";
        }
        else
        {
            ComputerName += L".html";
        }
        pszFileName = ComputerName.c_str();
    }
    else if ( tabIndex == TAB_INDEX_HELP )
    {
        pszFilter   = L"Rich Text (*.rtf)\0*.rtf\0\0";
        filterIndex = 1;
        pszFileName = L"WinAudit_Help.rtf";
    }
    else
    {
        pszFilter   = L"Text (*.txt)\0\0";
        filterIndex = 1;
        if ( tabIndex == TAB_INDEX_DISKS )
        {
            pszFileName = L"disk_information.txt";
        }
        else if ( tabIndex == TAB_INDEX_DISPLAYS )
        {
            pszFileName = L"display_information.txt";
        }
        else if ( tabIndex == TAB_INDEX_FIRMWARE )
        {
            pszFileName = L"firmware_information.txt";
        }
        else if ( tabIndex == TAB_INDEX_PROCESSORS )
        {
            pszFileName = L"processor_information.txt";
        }
        else if ( tabIndex == TAB_INDEX_SOFTWARE )
        {
            pszFileName = L"software_information.txt";
        }
        else if ( tabIndex == TAB_INDEX_LOGGER )
        {
            pszFileName = L"log.txt";
        }
        else
        {
            ErrorMessage  = L"tabIndex = ";
            ErrorMessage += Format.SizeT( tabIndex );
            throw SystemException( ERROR_INVALID_DATA, ErrorMessage.c_str(), __FUNCTION__ );
        }
    }

    if ( false == PXSPromptForFilePath( m_hWindow,
                                        false,
                                        true,
                                        pszFilter,
                                        &filterIndex, pszFileName, pszExtensions, &FilePath ) )
    {
        return;     // User cancelled
    }
    SaveToFile( tabIndex, filterIndex, FilePath );
    m_uSaveAuditFilterIndex = filterIndex;
}

//===============================================================================================//
//  Description:
//     Write the contents of the specified tab to a file
//
//  Parameters:
//      tabIndex   - the tab
//      dataFormat - audit data format: CSV, RTF or HTML
//      FilePath   - the output file path
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::SaveToFile( size_t tabIndex, DWORD dataFormat, const String& FilePath )
{
    char*    pszAnsi  = nullptr;
    DWORD    tableCounter = 0;
    size_t   numChars     = 0;
    File     FileObject;
    String   DataString, OutputFilePath, ExcelSepDeclaration;
    Window*  pWindow = nullptr;
    Formatter     Format;
    AllocateChars AnsiChars;
    TArray< TreeViewItem > CategoryItems;

    WaitCursor Wait;
    OutputFilePath = FilePath;
    OutputFilePath.Trim();

    if ( tabIndex == TAB_INDEX_AUDIT )
    {
        // On success, filterIndex has the user's choice of file type
        if ( dataFormat == DATA_FORMAT_CSV )
        {
            if ( OutputFilePath.EndsWithStringI( L".csv" ) == false )
            {
                OutputFilePath += L".csv";
            }
            PXSAuditRecordsToCsv( m_AuditRecords, false, &DataString );
            FileObject.CreateNew( OutputFilePath, 0, true );     // Unicode
            ExcelSepDeclaration = L"sep=,\r\n";
            FileObject.WriteChars( ExcelSepDeclaration );
            FileObject.WriteChars( DataString );
        }
        else if ( dataFormat == DATA_FORMAT_CSV2 )
        {
            if ( OutputFilePath.EndsWithStringI( L".csv2" ) == false )
            {
                OutputFilePath += L".csv2";
            }
            PXSAuditRecordsToCsv2( m_AuditRecords, true, &DataString );
            FileObject.CreateNew( OutputFilePath, 0, true );     // Unicode
            ExcelSepDeclaration = L"sep=,\r\n";
            FileObject.WriteChars( ExcelSepDeclaration );
            FileObject.WriteChars( DataString );
        }
        else if ( dataFormat == DATA_FORMAT_RTF )
        {
            // RTF, will recreate the rich text rather than use the window's
            // content in case need to verify the rich text is being correctly
            // created.
            if ( OutputFilePath.EndsWithStringI( L".rtf" ) == false )
            {
                OutputFilePath += L".rtf";
            }
            PXSAuditRecordsToContent( m_AuditRecords, &CategoryItems, &tableCounter, &DataString );
            numChars = DataString.GetAnsiMultiByteLength();
            pszAnsi  = AnsiChars.New( numChars );
            Format.StringToAnsi( DataString, pszAnsi, numChars );
            FileObject.CreateNew( OutputFilePath, 0, false );
            FileObject.Write( pszAnsi, numChars );
        }
        else
        {
            // HTML
            if ( ( OutputFilePath.EndsWithStringI( L".htm"  ) == false ) &&
                 ( OutputFilePath.EndsWithStringI( L".html" ) == false )  )
            {
                OutputFilePath += L".html";
            }
            PXSAuditRecordsToHtml( m_AuditRecords, &DataString );
            numChars = DataString.GetAnsiMultiByteLength();
            pszAnsi  = AnsiChars.New( numChars );
            Format.StringToAnsi( DataString, pszAnsi, numChars );
            FileObject.CreateNew( OutputFilePath, 0, false );
            FileObject.Write( pszAnsi, numChars );
        }
    }
    else if ( tabIndex == TAB_INDEX_HELP )
    {
        // Rich text
        if ( OutputFilePath.EndsWithStringI( L".rtf" ) == false )
        {
            OutputFilePath += L".rtf";
        }
        if ( g_pApplication == nullptr )
        {
            throw NullException( L"g_pApplication", __FUNCTION__ );
        }
        g_pApplication->LoadTextDataResource( IDR_WINAUDIT_HELP_RTF, &DataString );
        numChars = DataString.GetAnsiMultiByteLength();
        pszAnsi  = AnsiChars.New( numChars );
        Format.StringToAnsi( DataString, pszAnsi, numChars );
        FileObject.CreateNew( OutputFilePath, 0, false );
        FileObject.Write( pszAnsi, numChars );
    }
    else
    {
        // Plain text
        if ( OutputFilePath.EndsWithStringI( L".txt" ) == false )
        {
            OutputFilePath += L".txt";
        }

        pWindow = m_TabWindow.GetWindow( tabIndex );
        if ( pWindow == nullptr )
        {
            throw NullException( L"pWindow", __FUNCTION__ );
        }
        pWindow->GetText( &DataString );
        FileObject.CreateNew( OutputFilePath, 0, true );
        FileObject.WriteChars( DataString );
    }
}

//===============================================================================================//
//  Description:
//     Select the content of the selected tab window
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::SelectAllContents()
{
    String    ClassName;
    Window*   pWindow = m_TabWindow.GetSelectedTabWindow();
    TextArea* pTextArea;

    if ( pWindow == nullptr )
    {
       return;  // No tab selected so nothing to do
    }

    // Verify the window is an edit box or descendant
    pWindow->GetWndClassExClassName( &ClassName );
    if ( ( ClassName.CompareI( L"EDIT"   ) ) &&
         ( ClassName.CompareI( MSFTEDIT_CLASS ) )  )
    {
        throw SystemException( ERROR_INVALID_WINDOW_STYLE, L"EDIT", __FUNCTION__ );
    }
    pTextArea = static_cast< TextArea* >( pWindow );
    pTextArea->SelectAll();
}

//===============================================================================================//
//  Description:
//     Set the font into the currently selected tab
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::SelectFont()
{
    Font    FontObject;
    Window* pWindow = m_TabWindow.GetSelectedTabWindow();

    if ( pWindow == nullptr )
    {
       return;  // No tab selected so nothing to do
    }
    FontObject = pWindow->GetFont();

    if ( PXSShowChooseFontDialog( m_hWindow, &FontObject ) == false )
    {
        return;     // Operation cancelled
    }
    pWindow->SetFont( FontObject );
}

//===============================================================================================//
//  Description:
//     Set the main frame's bounds
//
//  Parameters:
//      ApplicationBounds - string of top,left,right,bottom values in pixels
//
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::SetApplicationBounds( const String& ApplicationBounds )
{
    RECT   bounds   = { 0, 0, 0, 0 };
    RECT   workArea = { 0, 0, 0, 0 };
    String Value;
    StringArray Values;

    SystemParametersInfo( SPI_GETWORKAREA, 0, &workArea, 0 );
    ApplicationBounds.ToArray( ',', &Values );
    if ( 4 == Values.GetSize() )
    {
        Value = Values.Get( 0 );
        if ( Value.c_str() )
        {
            bounds.left = _wtol( Value.c_str() );
            bounds.left = PXSMaxInt( 0, bounds.left );
        }

        Value = Values.Get( 1 );
        if ( Value.c_str() )
        {
            bounds.top = _wtol( Value.c_str() );
            bounds.top = PXSMaxInt( 0, bounds.top );
        }

        Value = Values.Get( 2 );
        if ( Value.c_str() )
        {
            bounds.right = _wtol( Value.c_str() );
            bounds.right = PXSMinInt( workArea.right, bounds.right );
        }

        Value = Values.Get( 3 );
        if ( Value.c_str() )
        {
            bounds.bottom = _wtol( Value.c_str() );
            bounds.bottom = PXSMinInt( workArea.bottom, bounds.bottom );
        }
    }

    // Verify its normal
    if ( ( bounds.right > bounds.left ) && ( bounds.bottom > bounds.top ) )
    {
        MoveWindow( m_hWindow,
                    bounds.left,
                    bounds.top, bounds.right - bounds.left, bounds.bottom - bounds.top, TRUE );
    }
    else
    {
        ShowWindow( m_hWindow, SW_MAXIMIZE );
    }
}

//===============================================================================================//
//  Description:
//     Adjust the size of the font for the selected tab window
//
//  Parameters:
//      bigger - true for a bigger size, false for smaller
//
//  Remarks:
//     The font should be scalable as adjusting the logical size
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::SetBiggerSmallerText( bool bigger )
{
    Font    FontObject;
    Window* pWindow = m_TabWindow.GetSelectedTabWindow();

    if ( pWindow == nullptr )
    {
       return;  // No tab selected so nothing to do
    }
    FontObject = pWindow->GetFont();

    // Increment or decrement
    if ( bigger )
    {
        FontObject.IncrementLogicalHeight();
    }
    else
    {
        FontObject.DecrementLogicalHeight();
    }
    FontObject.Create();
    pWindow->SetFont( FontObject );
}

//===============================================================================================//
//  Description:
//     Set the labels/text for the controls
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::SetControlLabels()
{
    String Caption, Find, FindBackwards, FindForwards, MatchCase;
    String TextNotFound, ToolTip, CollapseAll, ExpandAll, SelectAll;
    String DeselectAll;

    ///////////////////////////////////////////////////////////////////////////
    // Tool Bar

    // Audit Button
    PXSGetResourceString( PXS_IDS_1010_AUDIT, &Caption );
    m_AuditButton.SetText( Caption );
    PXSGetResourceString( PXS_IDS_1060_AUDIT_YOUR_COMPUTER, &ToolTip );
    m_AuditButton.SetToolTipText( ToolTip );

    // Stop Button
    PXSGetResourceString( PXS_IDS_1061_STOP, &Caption );
    m_StopButton.SetText( Caption );
    PXSGetResourceString( PXS_IDS_1062_STOP_THE_AUDIT, &ToolTip );
    m_StopButton.SetToolTipText( ToolTip );

    // Options Button
    PXSGetResourceString( PXS_IDS_1032_OPTIONS, &Caption );
    m_OptionsButton.SetText( Caption );
    PXSGetResourceString( PXS_IDS_1063_AUDIT_OPTIONS, &ToolTip );
    m_OptionsButton.SetToolTipText( ToolTip );

    // Save Button
    PXSGetResourceString( PXS_IDS_1011_SAVE, &Caption );
    m_SaveButton.SetText( Caption );
    PXSGetResourceString( PXS_IDS_1064_SAVE_TO_FILE, &ToolTip );
    m_SaveButton.SetToolTipText( ToolTip );

    // Email Button
    PXSGetResourceString( PXS_IDS_132_EMAIL, &Caption );
    m_EmailButton.SetText( Caption );
    PXSGetResourceString( PXS_IDS_1012_SEND_EMAIL, &ToolTip );
    m_EmailButton.SetToolTipText( ToolTip );

    // Help Button
    PXSGetResourceString( PXS_IDS_1065_HELP, &Caption );
    m_HelpButton.SetText( Caption );
    PXSGetResourceString( PXS_IDS_1066_VIEW_HELP, &ToolTip );
    m_HelpButton.SetToolTipText( ToolTip );

    ///////////////////////////////////////////////////////////////////////////
    // Find Text Bar

    PXSGetResourceString( PXS_IDS_500_FIND_COLON, &Find );
    PXSGetResourceString( PXS_IDS_501_FIND_BACKWARDS, &FindBackwards );
    PXSGetResourceString( PXS_IDS_502_FIND_FORWARDS, &FindForwards );
    PXSGetResourceString( PXS_IDS_503_MATCH_CASE, &MatchCase );
    PXSGetResourceString( PXS_IDS_504_TEXT_NOT_FOUND, &TextNotFound );
    m_FindTextBar.SetCaptions( Find, FindBackwards, FindForwards, MatchCase, TextNotFound );

    ///////////////////////////////////////////////////////////////////////////
    // Windows

    PXSGetResourceString( PXS_IDS_1268_COLLAPSE_ALL, &CollapseAll );
    PXSGetResourceString( PXS_IDS_1269_EXPAND_ALL  , &ExpandAll );
    PXSGetResourceString( PXS_IDS_1270_SELECT_ALL  , &SelectAll );
    PXSGetResourceString( PXS_IDS_1271_DESELECT_ALL, &DeselectAll );

    // Categories tree
    PXSGetResourceString( PXS_IDS_1160_CATEGORIES, &Caption );
    m_AuditCategories.SetClientTitleBarText( Caption );
    m_AuditCategories.SetMenuLabels( CollapseAll, ExpandAll, SelectAll, DeselectAll );

    // Help tree
    PXSGetResourceString( PXS_IDS_1065_HELP, &Caption );
    m_HelpContents.SetClientTitleBarText( Caption );
    m_HelpContents.SetMenuLabels( CollapseAll, ExpandAll, SelectAll, DeselectAll );

    // Tab Windows
    PXSGetResourceString( PXS_IDS_1010_AUDIT, &Caption );
    PXSGetResourceString( PXS_IDS_1210_AUDIT_REPORT_COMPUTER, &ToolTip );
    m_TabWindow.SetNameAndToolTip( TAB_INDEX_AUDIT, Caption, ToolTip );

    PXSGetResourceString( PXS_IDS_1211_DISKS, &Caption );
    PXSGetResourceString( PXS_IDS_1033_DISK_INFORMATION, &ToolTip );
    m_TabWindow.SetNameAndToolTip( TAB_INDEX_DISKS, Caption, ToolTip );

    PXSGetResourceString( PXS_IDS_1212_DISPLAYS, &Caption );
    PXSGetResourceString( PXS_IDS_1034_DISPLAY_INFORMATION, &ToolTip );
    m_TabWindow.SetNameAndToolTip( TAB_INDEX_DISPLAYS, Caption, ToolTip );

    PXSGetResourceString( PXS_IDS_1213_FIRMWARE, &Caption );
    PXSGetResourceString( PXS_IDS_1035_FIRMWARE_INFORMATION, &ToolTip );
    m_TabWindow.SetNameAndToolTip( TAB_INDEX_FIRMWARE, Caption, ToolTip );

    PXSGetResourceString( PXS_IDS_1140_PROCESSORS, &Caption );
    PXSGetResourceString( PXS_IDS_1037_PROCESSOR_INFORMATION,
                                     &ToolTip );
    m_TabWindow.SetNameAndToolTip( TAB_INDEX_PROCESSORS, Caption, ToolTip );

    PXSGetResourceString( PXS_IDS_1215_SOFTWARE, &Caption );
    PXSGetResourceString( PXS_IDS_1038_SOFTWARE_INFORMATION, &ToolTip );
    m_TabWindow.SetNameAndToolTip( TAB_INDEX_SOFTWARE, Caption, ToolTip );

    PXSGetResourceString( PXS_IDS_1065_HELP, &Caption );
    PXSGetResourceString( PXS_IDS_1066_VIEW_HELP, &ToolTip );
    m_TabWindow.SetNameAndToolTip( TAB_INDEX_HELP, Caption, ToolTip );

    PXSGetResourceString( PXS_IDS_1216_LOG, &Caption );
    PXSGetResourceString( PXS_IDS_1217_LOGGER_MESSAGES, &ToolTip );
    m_TabWindow.SetNameAndToolTip( TAB_INDEX_LOGGER, Caption, ToolTip );
}

//===============================================================================================//
//  Description:
//     Create an e-mail message of the contents of the currently selected tab
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::SendEmail()
{
    Mail      MailObject;
    File      FileObject;
    char*     pszAnsi  = nullptr;
    size_t    numChars = 0;
    size_t    tabIndex = m_TabWindow.GetSelectedTabIndex();
    Window*   pWindow  = nullptr;
    String    ErrorMessage, Address, Subject, NoteText, AttachFilePath;
    String    DataString;
    Formatter Format;
    Directory DirObject;
    AllocateChars     AnsiChars;
    SystemInformation SystemInfo;


    WaitCursor Wait;
    RedrawAllNow();
    if ( tabIndex == TAB_INDEX_AUDIT )
    {
        // Want content otherwise Outlook Express may put the html in
        // the message body
        NoteText = L"Computer Audit";
        PXSAuditRecordsToHtml( m_AuditRecords, &DataString );
        SystemInfo.GetComputerNetBiosName( &Subject );
        numChars = DataString.GetAnsiMultiByteLength();
        pszAnsi  = AnsiChars.New( numChars );
        Format.StringToAnsi( DataString, pszAnsi, numChars );
        DirObject.GetTempDirectory( &AttachFilePath );
        AttachFilePath += Subject;
        AttachFilePath += L".html";
        FileObject.CreateNew( AttachFilePath, 0, false );
        FileObject.Write( pszAnsi, numChars );
        FileObject.Close();
    }
    else if ( tabIndex == TAB_INDEX_HELP )
    {
        // Rich text
        if ( g_pApplication == nullptr )
        {
            throw NullException( L"g_pApplication", __FUNCTION__ );
        }
        g_pApplication->LoadTextDataResource( IDR_WINAUDIT_HELP_RTF, &NoteText );
        Subject = L"WinAudit Help";
    }
    else
    {
        // Plain text
        pWindow = m_TabWindow.GetWindow( tabIndex );
        if ( pWindow == nullptr )
        {
            throw NullException( L"pWindow", __FUNCTION__ );
        }
        pWindow->GetText( &NoteText );

        if ( tabIndex == TAB_INDEX_DISKS )
        {
            Subject = L"WinAudit Disk Information";
        }
        else if ( tabIndex == TAB_INDEX_DISPLAYS )
        {
            Subject = L"WinAudit Display Information";
        }
        else if ( tabIndex == TAB_INDEX_FIRMWARE )
        {
            Subject = L"WinAudit Firmware Information";
        }
        else if ( tabIndex == TAB_INDEX_PROCESSORS )
        {
            Subject = L"WinAudit Processor Information";
        }
        else if ( tabIndex == TAB_INDEX_SOFTWARE )
        {
            Subject = L"WinAudit Software Information";
        }
        else if ( tabIndex == TAB_INDEX_LOGGER )
        {
            Subject = L"WinAudit Log File";
        }
        else
        {
            ErrorMessage  = L"tabIndex = ";
            ErrorMessage += Format.SizeT( tabIndex );
            throw SystemException( ERROR_INVALID_DATA, ErrorMessage.c_str(), __FUNCTION__ );
        }
    }

    try
    {
        MailObject.SendSimpleMail( Address, Subject, NoteText, AttachFilePath );
        if ( AttachFilePath.GetLength() )
        {
           FileObject.Delete( AttachFilePath );
        }
    }
    catch ( const Exception& )
    {
        if ( AttachFilePath.GetLength() )
        {
            FileObject.Delete( AttachFilePath );
        }
        throw;
    }
}

//===============================================================================================//
//  Description:
//     Set the captions for the menu items
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::SetMenuLabels()
{
    String Label;

    // File Menu
    PXSGetResourceString( PXS_IDS_1000_FILE, &Label );
    m_FileMenuBarItem.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1010_AUDIT, &Label );
    Label += L"\tCtrl+N";
    m_FileAudit.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1011_SAVE, &Label );
    Label += L"\tCtrl+S";
    m_FileSave.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1012_SEND_EMAIL, &Label );
    m_FileEmail.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1013_DATABASE_EXPORT, &Label );
    m_FileOdbc.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1014_DESKTOP_SHORTCUT, &Label );
    m_FileShortcut.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_133_EXIT, &Label );
    Label += L"\tAlt+F4";
    m_FileExit.SetLabel( Label );

    // Edit Menu
    PXSGetResourceString( PXS_IDS_1001_EDIT, &Label );
    m_EditMenuBarItem.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_131_SELECT_ALL, &Label );
    Label += L"\tCtrl+A";
    m_EditSelectAll.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_129_COPY, &Label );
    Label += L"\tCtrl+C";
    m_EditCopy.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1020_FIND, &Label );
    Label += L"\tCtrl+F";
    m_EditFind.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1021_FONT, &Label );
    m_EditSetFont.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1022_SMALLER_TEXT, &Label );
    m_EditSmallerText.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1023_BIGGER_TEXT, &Label );
    m_EditBiggerText.SetLabel( Label );

    // View Menu
    PXSGetResourceString( PXS_IDS_1002_VIEW, &Label );
    m_ViewMenuBarItem.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1030_TOOLBAR, &Label );
    m_ViewToolBar.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1031_STATUSBAR, &Label );
    m_ViewStatusBar.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1032_OPTIONS, &Label );
    Label += L"\tCtrl+O";
    m_ViewOptions.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1037_PROCESSOR_INFORMATION, &Label );
    m_ViewProcessorInfo.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1033_DISK_INFORMATION, &Label );
    m_ViewDiskInfo.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1034_DISPLAY_INFORMATION, &Label );
    m_ViewDisplayInfo.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1036_POLICY_INFORMATION, &Label );
    m_ViewPolicyInfo.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1038_SOFTWARE_INFORMATION, &Label );
    m_ViewSoftwareInfo.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1035_FIRMWARE_INFORMATION, &Label );
    m_ViewFirmwareInfo.SetLabel( Label );

    // Language Menu
    PXSGetResourceString( PXS_IDS_1003_LANGUAGE, &Label );
    m_LanguageMenuBarItem.SetLabel( Label );

    // Chinese traditional
    Label  = 0x6B63;
    Label += 0x9AD4;
    Label += 0x4E2D;
    Label += 0x6587;
    m_LanguageChineseTw.SetLabel( Label );

    // Czech
    Label  = 0x010C;      // Capital C with caron 0xC8 in Central Europe
    Label += L"e";
    Label += 0x0161;      // Small s with caron in Central Europe
    Label += L"tina";
    m_LanguageCzech.SetLabel( Label );

    // Danish
    Label = L"Dansk";
    m_LanguageDanish.SetLabel( Label );

    // German
    Label = L"Deutsch";
    m_LanguageGerman.SetLabel( Label );

    // English
    Label = L"English";
    m_LanguageEnglish.SetLabel( Label );

    // Spanish
    Label  = L"Espa";
    Label += 0x00F1;
    Label += L"ol";
    Label += 0x00ED;
    m_LanguageSpanish.SetLabel( Label );

    // Greek
    Label  = 0x0395;   // Epsilon
    Label += 0x03bb;   // Lambda
    Label += 0x03bb;   // Lambda
    Label += 0x03b7;   // Eta
    Label += 0x03bd;   // Nu
    Label += 0x03b9;   // Iota
    Label += 0x03ba;   // Kappa
    Label += 0x03ac;   // Alpha with tonos
    m_LanguageGreek.SetLabel( Label );

    // Fench - Belgian
    Label  = L"Fran";
    Label += 0x00E7;        // cedilla in the ANSI West European code page
    Label += L"ais - BE";
    m_LanguageBelgian.SetLabel( Label );

    // French - France
    Label  = L"Fran";
    Label += 0x00E7;          // cedilla in the ANSI West European code page
    Label += L"ais - FR";
    m_LanguageFrench.SetLabel( Label );

    // Hebrew
    Label  = 0x05e2;   // Ayin, 0x05e2
    Label += 0x05d1;   // Bet, 0x05d1
    Label += 0x05e8;   // Resh, 0x05e8
    Label += 0x05d9;   // Yod, 0x05d9
    Label += 0x05ea;   // Tav, 0x05ea
    m_LanguageHebrew.SetLabel( Label );

    // Italian
    Label = L"Italiano";
    m_LanguageItalian.SetLabel( Label );

    // Indonesian
    Label = L"Indonesian";
    m_LanguageIndonesian.SetLabel( Label );

    // Japanese
    Label  = 0xE565;
    Label += 0x672C;
    Label += 0x8A9E;
    m_LanguageJapanese.SetLabel( Label );

    // Korean
    Label  = 0xD55C;
    Label += 0xAD6D;
    Label += 0xC5B4;
    m_LanguageKorean.SetLabel( Label );

    // Hungarian
    Label = L"Magyar";
    m_LanguageHungarian.SetLabel( Label );

    // Dutch
    Label = L"Nederlands";
    m_LanguageDutch.SetLabel( Label );

    // Polish
    Label = L"Polski";
    m_LanguagePolish.SetLabel( Label );

    // Portuguese - Brazil
    Label  = L"Portugu";
    Label += 0x00EA;        // e with circumflex in the ANSI West Europe
    Label += L"s - BR";
    m_LanguageBrazilian.SetLabel( Label );

    // Portuguese
    Label  = L"Portugu";
    Label += 0x00EA;          // e with circumflex in the ANSI West Europe
    Label += L"s - PT";
    m_LanguagePortuguese.SetLabel( Label );

    // Russian
    Label  = 0x0420;
    Label += 0x0443;
    Label += 0x0441;
    Label += 0x0441;
    Label += 0x043A;
    Label += 0x0438;
    Label += 0x0439;
    m_LanguageRussian.SetLabel( Label );

    // Slovak
    Label = L"Slovensky";
    m_LanguageSlovak.SetLabel( Label );

    // Serbian
    Label = L"Srpski";
    m_LanguageSerbian.SetLabel( Label );

    // Finnish
    Label = L"Suomi";
    m_LanguageFinnish.SetLabel( Label );

    // Thai
    Label  = 0x0e20;
    Label += 0x0e32;
    Label += 0x0e29;
    Label += 0x0e32;
    Label += 0x0e44;
    Label += 0x0e17;
    Label += 0x0e22;
    m_LanguageThai.SetLabel( Label );

    // Turkish
    Label  = PXS_STRING_EMPTY;
    Label += L"T";
    Label += 0x00FC;
    Label += L"rk";
    Label += 0x00E7;
    m_LanguageTurkish.SetLabel( Label );

    // Help Menu
    PXSGetResourceString( PXS_IDS_1004_HELP, &Label );
    m_HelpMenuBarItem.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1050_USING_WINAUDIT, &Label );
    Label += L"\tF1";
    m_HelpUsingWinAudit.SetLabel( Label );

    if ( g_pApplication )
    {
        g_pApplication->GetWebSiteURL( &Label );
    }
    m_HelpWebsite.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1051_ABOUT_WINAUDIT, &Label );
    m_HelpAbout.SetLabel( Label );

    PXSGetResourceString( PXS_IDS_1052_CREATE_GUID, &Label );
    m_HelpCreateGuid.SetLabel( Label );

    // Log File
    PXSGetResourceString( PXS_IDS_1053_START_LOGGING, &Label );
    m_HelpLogFile.SetLabel( Label );
}

//===============================================================================================//
//  Description:
//     Set the report oprions string to the configuration
//
//  Parameters:
//      ReportOptions - the report options
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::SetReportOptions( const String& ReportOptions )
{
    wchar_t ch = 0;
    size_t  i  = 0;
    size_t  length = ReportOptions.GetLength();
    String  LogMessage;

    // Reset all options to false so will only do what is specified
    m_ConfigurationSettings.systemOverview    = false;
    m_ConfigurationSettings.installedSoftware = false;
    m_ConfigurationSettings.operatingSystem   = false;
    m_ConfigurationSettings.peripherals       = false;
    m_ConfigurationSettings.security          = false;
    m_ConfigurationSettings.groupsAndUsers    = false;
    m_ConfigurationSettings.scheduledTasks    = false;
    m_ConfigurationSettings.uptimeStatistics  = false;
    m_ConfigurationSettings.errorLogs         = false;
    m_ConfigurationSettings.environmentVars   = false;
    m_ConfigurationSettings.regionalSettings  = false;
    m_ConfigurationSettings.windowsNetwork    = false;
    m_ConfigurationSettings.networkTcpIp      = false;
    m_ConfigurationSettings.hardwareDevices   = false;
    m_ConfigurationSettings.displays          = false;
    m_ConfigurationSettings.displayAdapters   = false;
    m_ConfigurationSettings.installedPrinters = false;
    m_ConfigurationSettings.biosVersion       = false;
    m_ConfigurationSettings.systemManagement  = false;
    m_ConfigurationSettings.processors        = false;
    m_ConfigurationSettings.memory            = false;
    m_ConfigurationSettings.physicalDisks     = false;
    m_ConfigurationSettings.drives            = false;
    m_ConfigurationSettings.commPorts         = false;
    m_ConfigurationSettings.startupPrograms   = false;
    m_ConfigurationSettings.services          = false;
    m_ConfigurationSettings.runningPrograms   = false;
    m_ConfigurationSettings.odbcInformation   = false;
    m_ConfigurationSettings.oleDbProviders    = false;
    m_ConfigurationSettings.softwareMetering  = false;
    m_ConfigurationSettings.userLogonStats    = false;
    for ( i = 0; i < length; i++ )
    {
        ch = ReportOptions.CharAt( i );
        switch ( ch )
        {
            default:
                LogMessage  = L"Unrecognised report option '";
                LogMessage += ch;
                LogMessage += L"'.";
                PXSLogAppWarn( LogMessage.c_str() );
                break;

           case 'g':
               m_ConfigurationSettings.systemOverview = true;
               break;

           case 's':
               m_ConfigurationSettings.installedSoftware = true;
               break;

           case 'o':
               m_ConfigurationSettings.operatingSystem = true;
               break;

           case 'P':
               m_ConfigurationSettings.peripherals = true;
               break;

           case 'x':
               m_ConfigurationSettings.security = true;
               break;

           case 'u':
               m_ConfigurationSettings.groupsAndUsers = true;
               break;

           case 'T':
               m_ConfigurationSettings.scheduledTasks = true;
               break;

           case 'U':
               m_ConfigurationSettings.uptimeStatistics = true;
               break;

           case 'e':
               m_ConfigurationSettings.errorLogs = true;
               break;

           case 'E':
               m_ConfigurationSettings.environmentVars = true;
               break;

           case 'R':
               m_ConfigurationSettings.regionalSettings = true;
               break;

           case 'N':
               m_ConfigurationSettings.windowsNetwork = true;
               break;

           case 't':
               m_ConfigurationSettings.networkTcpIp = true;
               break;

           case 'z':
               m_ConfigurationSettings.hardwareDevices = true;
               break;

           case 'D':
               m_ConfigurationSettings.displays = true;
               break;

           case 'a':
               m_ConfigurationSettings.displayAdapters = true;
               break;

           case 'I':
               m_ConfigurationSettings.installedPrinters = true;
               break;

           case 'b':
               m_ConfigurationSettings.biosVersion = true;
               break;

           case 'M':
               m_ConfigurationSettings.systemManagement = true;
               break;

           case 'p':
               m_ConfigurationSettings.processors = true;
               break;

           case 'm':
               m_ConfigurationSettings.memory = true;
               break;

           case 'i':
               m_ConfigurationSettings.physicalDisks = true;
               break;

           case 'd':
               m_ConfigurationSettings.drives = true;
               break;

           case 'c':
               m_ConfigurationSettings.commPorts = true;
               break;

           case 'S':
               m_ConfigurationSettings.startupPrograms = true;
               break;

           case 'A':
               m_ConfigurationSettings.services = true;
               break;

           case 'r':
               m_ConfigurationSettings.runningPrograms = true;
               break;

           case 'C':
               m_ConfigurationSettings.odbcInformation  = true;
               break;

           case 'O':
               m_ConfigurationSettings.oleDbProviders = true;
               break;

           case 'H':
               m_ConfigurationSettings.softwareMetering = true;
               break;

           case 'G':
               m_ConfigurationSettings.userLogonStats = true;
               break;
        }
    }
}

//===============================================================================================//
//  Description:
//     Set the colours used by the application
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::SetColours()
{
    COLORREF gradientColour = CLR_INVALID;
    COLORREF fillColour     = PXS_COLOUR_LITEGREY;
    COLORREF textHiLiteBackground = GetSysColor( COLOR_HIGHLIGHT );

    if ( PXSIsHighColour( nullptr ) )
    {
        gradientColour = RGB( 232, 232, 232 );
        fillColour     = RGB( 128, 196, 255 );
        textHiLiteBackground = RGB( 127, 159, 191 );
    }

    // Application object
    if ( g_pApplication )
    {
        g_pApplication->SetGradientColours( PXS_COLOUR_WHITE, gradientColour );
    }

    // Menu items
    SetBackgroundGradient( PXS_COLOUR_WHITE, gradientColour, true );

    // Controls
    m_ToolBar.SetBackgroundGradient( PXS_COLOUR_WHITE, gradientColour, true );

    m_AuditButton.SetBackgroundGradient( PXS_COLOUR_WHITE,
                                         gradientColour, true );
    if ( IsRightToLeftReading() )
    {
        m_AuditButton.SetFilledBitmap( fillColour, PXS_COLOUR_BLACK, PXS_SHAPE_TRIANGLE_LEFT );
    }
    else
    {
        m_AuditButton.SetFilledBitmap( fillColour, PXS_COLOUR_BLACK, PXS_SHAPE_TRIANGLE_RIGHT );
    }

    m_StopButton.SetBackgroundGradient( PXS_COLOUR_WHITE, gradientColour, true);
    m_StopButton.SetFilledBitmap( PXS_COLOUR_RED, PXS_COLOUR_BLACK, PXS_SHAPE_RECTANGLE );

    m_OptionsButton.SetBackgroundGradient( PXS_COLOUR_WHITE, gradientColour, true );
    m_OptionsButton.SetFilledBitmap( fillColour, PXS_COLOUR_BLACK, PXS_SHAPE_DIAMOND );

    m_SaveButton.SetBackgroundGradient( PXS_COLOUR_WHITE, gradientColour, true);
    m_SaveButton.SetFilledBitmap( fillColour, PXS_COLOUR_BLACK, PXS_SHAPE_RECTANGLE_HOLLOW);

    m_EmailButton.SetBackgroundGradient(PXS_COLOUR_WHITE, gradientColour, true);
    m_EmailButton.SetFilledBitmap( fillColour, PXS_COLOUR_BLACK, PXS_SHAPE_FIVE_SIDES );

    m_HelpButton.SetBackgroundGradient( PXS_COLOUR_WHITE, gradientColour, true);
    m_HelpButton.SetFilledBitmap( fillColour, PXS_COLOUR_BLACK, PXS_SHAPE_CROSS );

    m_FindTextBar.SetBackground( gradientColour );
    m_FindTextBar.SetControlColours( fillColour );

    m_SidePanel.SetBackgroundGradient( PXS_COLOUR_WHITE, gradientColour, false);

    m_AuditCategories.SetBackgroundGradient( PXS_COLOUR_WHITE, gradientColour, false);
    m_AuditCategories.SetTextHiLiteBackground( textHiLiteBackground );
    m_AuditCategories.SetClientTitleBarColour( fillColour );

    m_HelpContents.SetBackgroundGradient( PXS_COLOUR_WHITE, gradientColour, false );
    m_HelpContents.SetTextHiLiteBackground( textHiLiteBackground );
    m_HelpContents.SetClientTitleBarColour( fillColour );

    m_TabWindow.SetBackgroundGradient( PXS_COLOUR_WHITE, gradientColour, true );
    m_TabWindow.SetTabColours( PXS_COLOUR_WHITE, gradientColour, fillColour );

    m_Splitter.SetBackgroundGradient( PXS_COLOUR_WHITE, gradientColour, true );

    m_StatusBar.SetBackgroundGradient( PXS_COLOUR_WHITE, gradientColour, true);

    m_ProgressBar.SetForeground( fillColour );
    m_ProgressBar.SetBackgroundGradient(PXS_COLOUR_WHITE, gradientColour, true);
    Repaint();
}

//===============================================================================================//
//  Description:
//     Set the state of the items on the Edit Menu
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::SetEditMenu()
{
    bool    textSelected = false;
    String  ClassName;
    Window* pWindow = nullptr;
    RichEditBox* pRichEdit = nullptr;

    // Set the Edit menu for rich edit controls
    pWindow = m_TabWindow.GetSelectedTabWindow();
    if ( pWindow )
    {
        // Verify its a rich edit control
        pWindow->GetWndClassExClassName( &ClassName );
        if ( ClassName.CompareI( MSFTEDIT_CLASS ) == 0 )
        {
            pRichEdit    = static_cast< RichEditBox* >( pWindow );
            textSelected = pRichEdit->IsAnyTextSelected();
            m_EditCopy.SetEnabled( textSelected );
        }
    }
}

//===============================================================================================//
//  Description:
//     Show the about dialog
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::ShowAboutDialog()
{
    String      Title;
    AboutDialog Dialog;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }
    PXSGetResourceString( PXS_IDS_1051_ABOUT_WINAUDIT, &Title );
    Dialog.SetTitle( Title );
    Dialog.Create( m_hWindow );
}

//===============================================================================================//
//  Description:
//     Show the command line help
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::ShowCommandLineHelpMessage()
{
    String Help, ApplicationName;
    MessageDialog Dialog;

    Help.Allocate( 256 );
    Help  = L"WinAudit command line usage:";
    Help += PXS_STRING_CRLF;
    Help += PXS_STRING_CRLF;
    Help += L"WinAudit /h /r=report /f=file /l=log_file /T=timestamp";
    Help += PXS_STRING_CRLF;
    Help += PXS_STRING_CRLF;
    Help += L"/h\tShow this help message then exit";
    Help += PXS_STRING_CRLF;
    Help += L"/r\tReport content, categories are identified by case";
    Help += PXS_STRING_CRLF;
    Help += L"\tsensitive letters, see the help documentation for examples";
    Help += PXS_STRING_CRLF;
    Help += L"/f\tFile name for output or database connection string";
    Help += PXS_STRING_CRLF;
    Help += L"/l\tLog file path";
    Help += PXS_STRING_CRLF;
    Help += L"/T\tTimestamp in output file name";
    Help += PXS_STRING_CRLF;

    PXSGetApplicationName( &ApplicationName );
    Dialog.SetTitle( ApplicationName );
    Dialog.SetSize( 500, 250 );
    Dialog.SetMessage( Help );
    Dialog.Create( m_hWindow );
}

//===============================================================================================//
//  Description:
//     Create and show a GUID for registyr/database insertion
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::ShowCreateGuid()
{
    String  ApplicationName, Message, GuidString;
    Formatter     Format;
    MessageDialog GuidDialog;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }
    PXSGetApplicationName( &ApplicationName );
    GuidDialog.SetTitle( ApplicationName );

    GuidString = Format.CreateGuid();

    Message.Allocate( 1024 );
    Message  = L"\t\tGLOBALLY UNIQUE IDENTIFIER (GUID)\r\n";
    Message += PXS_STRING_CRLF;
    Message += PXS_STRING_CRLF;
    Message += L"When WinAudit posts records to a database, it identifies "
               L"known computers by\r\n"
               L"matching values found in the Computer_Master table. "
               L"Unfortunately some\r\n"
               L"computers have the same identifier values. You can "
               L"diagnose this by looking\r\n"
               L"in WinAudit's log file for the message \"Matched computer "
               L"using identifier...\".\r\n"
               L"If two computers are matched on the same identifier value"
               L" then WinAudit\r\n"
               L"considers them to be the same computer. To solve this, "
               L"follow the steps\r\n"
               L"below. The WinAuditGuid takes precedence over all other "
               L"identifiers.\r\n"
               L"Each time this window is shown a new GUID is created.\r\n"
               L"_________________________________________________________"
               L"\r\n\r\nSteps:\r\n\r\n"
               L"1. In the registry on the mis-identified computer, "
               L"create the path \r\n"
               L"    'HKEY_LOCAL_MACHINE\\";
    Message += PXS_REG_PATH_COMPUTER_GUID;
    Message += L"'";
    Message += PXS_STRING_CRLF;
    Message += PXS_STRING_CRLF;
    Message += L"2. Create the string value '";
    Message += PXS_WINAUDIT_COMPUTER_GUID;
    Message += L"'";
    Message += PXS_STRING_CRLF;
    Message += PXS_STRING_CRLF;
    Message += L"3. Put the GUID ";
    Message += GuidString;
    Message += L" into WinAuditGuid\r\n";
    Message += PXS_STRING_CRLF;
    Message += L"4. Put the same GUID ";
    Message += GuidString;
    Message += L" into the\r\n";
    Message += L"    database field \"Computer_Master.WinAudit_GUID\" of "
               L"the row that\r\n";
    Message += L"    identifies the computer\r\n";
    Message += PXS_STRING_CRLF;
    Message += L"5. For each mis-identified computer, close then re-show "
               L"this window to\r\n";
    Message += L"    create a new GUID. Repeat steps 1 to 4.\r\n";
    Message += PXS_STRING_CRLF;
    Message += L"_________________________________________________________";
    Message += PXS_STRING_CRLF;
    Message += PXS_STRING_CRLF;
    GuidDialog.SetMessage( Message );
    GuidDialog.SetSize( 600, 400 );
    GuidDialog.Create( m_hWindow );
}

//===============================================================================================//
//  Description:
//      Show/hide the help
//
//  Parameters:
//      show - true if want to show the side panel, else false
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::ShowHideHelp( bool show )
{
    WaitCursor Wait;

    if ( show )
    {
        m_AuditCategories.SetVisible( false );
        m_HelpContents.SetVisible( true );
        m_HelpContents.ExpandOrCollapseAll( true );
        ShowHideSidePanel( true );
        m_TabWindow.SetSelectedTabIndex( TAB_INDEX_HELP );
        m_TabWindow.SetTabVisible( TAB_INDEX_HELP, true );
    }
    else
    {
        m_AuditCategories.SetVisible( true );
        m_HelpContents.SetVisible( false );
        m_TabWindow.SetSelectedTabIndex( TAB_INDEX_AUDIT );
        m_TabWindow.SetTabVisible( TAB_INDEX_HELP, false );
    }
    DoLayout();
}

//===============================================================================================//
//  Description:
//      Resize the side panel to show/hide it by setting its width
//
//  Parameters:
//      show - true if want to show the side panel, else false
//
//  Remarks:
//      Does not change the panel's width if it is already visible
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::ShowHideSidePanel( bool show )
{
    if ( show )
    {
        // Ensure has a non-zero width. Will use 140 pixels as this allows
        // the data tables to be fully seen at 800x600 pixel monitor
        if ( IsRightToLeftReading() )
        {
            if ( m_Splitter.GetOffsetTwo() == 0 )
            {
                m_Splitter.SetOffsetTwo( 140 );
            }
        }
        else
        {
            if ( m_Splitter.GetOffset() == 0 )
            {
                m_Splitter.SetOffset( 140 );
            }
        }
    }
    else
    {
        // Shrink to zero width
        if ( IsRightToLeftReading() )
        {
            m_Splitter.SetOffsetTwo( 0 );
        }
        else
        {
            m_Splitter.SetOffset( 0 );
        }
    }
    DoLayout();
}

//===============================================================================================//
//  Description:
//     Show the dialog box to export the records to a database
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::ShowOdbcExport()
{
    String  ApplicationName;
    AuditData   Auditor;
    AuditRecord AuditMasterRecord, ComputerMasterRecord;
    OdbcExportDialog ExportDialog;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    WaitCursor Wait;
    RedrawAllNow();
    Auditor.MakeAuditMasterRecord( &AuditMasterRecord );
    Auditor.MakeComputerMasterRecord( &ComputerMasterRecord );
    PXSGetApplicationName( &ApplicationName );
    ExportDialog.SetTitle( ApplicationName );
    ExportDialog.SetAuditRecords( AuditMasterRecord, ComputerMasterRecord, m_AuditRecords );
    ExportDialog.SetConfigurationSettings( m_ConfigurationSettings );
    Wait.Restore();
    if ( IDOK == ExportDialog.Create( m_hWindow ) )
    {
        ExportDialog.GetConfigurationSettings( &m_ConfigurationSettings );
    }
}

//===============================================================================================//
//  Description:
//     Show the CPUID data
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::ViewDiskInformation()
{
    String          DiskDiagnostics, ApplicationDiagnostics;
    DiskInformation DiskInfo;

    if ( m_bCreatedControls == false )
    {
        throw FunctionException( L"m_bCreatedControls", __FUNCTION__ );
    }
    WaitCursor Wait;
    ShowHideSidePanel( false );
    m_TabWindow.SetSelectedTabIndex( TAB_INDEX_DISKS );
    m_TabWindow.SetTabVisible( TAB_INDEX_DISKS, true );
    DiskInfo.GetDiagnostics( &DiskDiagnostics );
    if ( g_pApplication )
    {
        g_pApplication->GetDiagnostics( &ApplicationDiagnostics );
    }
    DiskDiagnostics += PXS_STRING_CRLF;
    DiskDiagnostics += ApplicationDiagnostics;
    m_DiskRichBox.SetText( DiskDiagnostics );
}

//===============================================================================================//
//  Description:
//     Show the display data
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::ViewDisplayInformation()
{
    String DisplayDiagnostics, ApplicationDiagnostics;
    DisplayInformation DisplayInfo;


    if ( m_bCreatedControls == false )
    {
        throw FunctionException( L"m_bCreatedControls", __FUNCTION__ );
    }
    WaitCursor Wait;
    ShowHideSidePanel( false );
    m_TabWindow.SetSelectedTabIndex( TAB_INDEX_DISPLAYS );
    m_TabWindow.SetTabVisible( TAB_INDEX_DISPLAYS, true );
    DisplayInfo.GetDiagnostics( &DisplayDiagnostics );
    if ( g_pApplication )
    {
        g_pApplication->GetDiagnostics( &ApplicationDiagnostics );
    }
    DisplayDiagnostics += PXS_STRING_CRLF;
    DisplayDiagnostics += ApplicationDiagnostics;
    m_DisplayRichBox.SetText( DisplayDiagnostics );
}

//===============================================================================================//
//  Description:
//     Show the system's firmware
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::ViewFirmwareInformation()
{
    String FirmwareData, FirmwareDataNT6, ApplicationDiagnostics;
    WindowsInformation WindowsInfo;

    if ( m_bCreatedControls == false )
    {
        throw FunctionException( L"m_bCreatedControls", __FUNCTION__ );
    }

    if ( g_pApplication == nullptr )
    {
        throw NullException( L"g_pApplication", __FUNCTION__ );
    }
    WaitCursor Wait;
    ShowHideSidePanel( false );
    m_TabWindow.SetSelectedTabIndex( TAB_INDEX_FIRMWARE );
    m_TabWindow.SetTabVisible( TAB_INDEX_FIRMWARE, true );

    // XP and newer
    WindowsInfo.GetFirmwareDataXP( &FirmwareData );

    // NT6 and newer
    if ( WindowsInfo.GetMajorVersion() >= 6 )
    {
        FirmwareData += PXS_STRING_CRLF;
        FirmwareData += PXS_STRING_CRLF;
        FirmwareData += PXS_STRING_CRLF;
        FirmwareData += PXS_STRING_CRLF;
        WindowsInfo.GetFirmwareDataNT6( &FirmwareDataNT6 );
        FirmwareData += FirmwareDataNT6;
        FirmwareData += PXS_STRING_CRLF;
    }
    FirmwareData += PXS_STRING_CRLF;
    FirmwareData += PXS_STRING_CRLF;
    FirmwareData += PXS_STRING_CRLF;
    if ( g_pApplication )
    {
        g_pApplication->GetDiagnostics( &ApplicationDiagnostics );
    }
    FirmwareData += ApplicationDiagnostics;
    m_FirmwareRichBox.SetText( FirmwareData );
}

//===============================================================================================//
//  Description:
//     Show the audit options dialog
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::ViewOptionsDialog()
{
    String  ApplicationName;
    WinAuditConfigDialog ConfigDialog;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }
    PXSGetApplicationName( &ApplicationName );
    ConfigDialog.SetTitle( ApplicationName );
    ConfigDialog.SetConfigurationSettings( m_ConfigurationSettings );
    if ( IDOK == ConfigDialog.Create( m_hWindow ) )
    {
        ConfigDialog.GetConfigurationSettings( &m_ConfigurationSettings );
        AuditStart();
    }
}

//===============================================================================================//
//  Description:
//     Show the resultant set of policy data in the browser
//
//  Parameters:
//      none
//
//  Remarks:
//      Vista or newer. Data is in HTML format so will show in the browser
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::ViewPolicyInformation()
{
    File   FileObject;
    Shell  ShellObject;
    String RsopHtmlResult, FilePath;
    SecurityInformation SecurityInfo;

    WaitCursor Wait;
    SecurityInfo.GetRsopHtmlResult( &RsopHtmlResult );
    if ( g_pApplication )
    {
        g_pApplication->GetApplicationTempPath( false, &FilePath );
    }
    FilePath += L"rsop.html";
    FileObject.CreateNew( FilePath, 0, false );
    FileObject.WriteChars( RsopHtmlResult );
    ShellObject.NavigateToUrl( FilePath );
}

//===============================================================================================//
//  Description:
//     Show the CPUID data
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::ViewProcessorInformation()
{
    String         Diagnostics, ApplicationDiagnostics;
    CpuInformation CpuInfo;

    if ( m_bCreatedControls == false )
    {
        throw FunctionException( L"m_bCreatedControls", __FUNCTION__ );
    }
    WaitCursor Wait;
    ShowHideSidePanel( false );
    m_TabWindow.SetSelectedTabIndex( TAB_INDEX_PROCESSORS );
    m_TabWindow.SetTabVisible( TAB_INDEX_PROCESSORS, true );
    CpuInfo.GetDiagnostics( &Diagnostics );
    if ( g_pApplication )
    {
        g_pApplication->GetDiagnostics( &ApplicationDiagnostics );
    }
    Diagnostics += PXS_STRING_CRLF;
    Diagnostics += ApplicationDiagnostics;
    m_ProcessorRichBox.SetText( Diagnostics );
}

//===============================================================================================//
//  Description:
//     Show the software data
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::ViewSoftwareInformation()
{
    String SoftwareDiagnostics, ApplicationDiagnostics;
    SoftwareInformation SoftwareInfo;

    if ( m_bCreatedControls == false )
    {
        throw FunctionException( L"m_bCreatedControls", __FUNCTION__ );
    }
    WaitCursor Wait;
    ShowHideSidePanel( false );
    m_TabWindow.SetSelectedTabIndex( TAB_INDEX_SOFTWARE );
    m_TabWindow.SetTabVisible( TAB_INDEX_SOFTWARE, true );
    SoftwareInfo.GetDiagnostics( &SoftwareDiagnostics );
    if ( g_pApplication )
    {
        g_pApplication->GetDiagnostics( &ApplicationDiagnostics );
    }
    SoftwareDiagnostics += PXS_STRING_CRLF;
    SoftwareDiagnostics += ApplicationDiagnostics;
    m_SoftwareRichBox.SetText( SoftwareDiagnostics );
}

//===============================================================================================//
//  Description:
//     Start/stop the logger with the messages going to a string array
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::StartStopLogger( bool start )
{
    if ( m_bCreatedControls == false )
    {
        throw FunctionException( L"m_bCreatedControls", __FUNCTION__ );
    }

    if ( g_pApplication == nullptr )
    {
        throw NullException( L"g_pApplication", __FUNCTION__ );
    }

    if ( start )
    {
        // Set to maximum log messages
        g_pApplication->SetLogLevel( PXS_LOG_LEVEL_VERBOSE );
        g_pApplication->StartLogger( false, true );
        m_TabWindow.SetTabVisible( TAB_INDEX_LOGGER, true );
        m_HelpLogFile.SetChecked( true );
    }
    else
    {
        // Will leave the tab visible
        g_pApplication->StopLogger();
        m_HelpLogFile.SetChecked( false );
    }
    SetControlsAndMenus();
    UpdateLogWindow();
}

//===============================================================================================//
//  Description:
//     Update the log window with any messages in the logger
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::UpdateLogWindow()
{
    size_t i = 0, numMessages = 0;
    String Line;
    StringArray LogMessages;

    if ( g_pApplication == nullptr )
    {
        throw NullException( L"g_pApplication", __FUNCTION__ );
    }

    // Suppress all errors to avoid log recursion
    try
    {
        g_pApplication->GetLogMessages( true, &LogMessages );
        numMessages = LogMessages.GetSize();
        for ( i = 0; i < numMessages; i++ )
        {
            Line = LogMessages.Get( i );
            Line += PXS_STRING_CRLF;
            m_LogRichBox.AppendText( Line );
        }
    }
    catch ( const Exception& )
    { }
}

//===============================================================================================//
//  Description:
//     Write out the ini file
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void WinAuditFrame::WriteIniFile()
{
    RECT bounds = { 0, 0, 0, 0 };
    File FileObject;
    String Content, IniPath;
    Formatter Format;

    Content.Allocate( 1024 );

    // Header
    Content  = L"!Parmavex WinAudit Freeware Configuration File\r\n"
               L"!You can delete this file at any time. If an ini\r\n"
               L"!file is not present in the same directory as the \r\n"
               L"!executable, the default settings will be used.\r\n";
    Content += PXS_STRING_CRLF;

    // Auto-start
    Content += L"autoStartAudit=";
    if ( m_ConfigurationSettings.autoStartAudit )
    {
        Content += L"1";
    }
    else
    {
        Content += L"0";
    }
    Content += PXS_STRING_CRLF;

    // Report
    Content += L"r=";
    if ( m_ConfigurationSettings.systemOverview    ) Content += 'g';
    if ( m_ConfigurationSettings.installedSoftware ) Content += 's';
    if ( m_ConfigurationSettings.operatingSystem   ) Content += 'o';
    if ( m_ConfigurationSettings.peripherals       ) Content += 'P';
    if ( m_ConfigurationSettings.security          ) Content += 'x';
    if ( m_ConfigurationSettings.groupsAndUsers    ) Content += 'u';
    if ( m_ConfigurationSettings.scheduledTasks    ) Content += 'T';
    if ( m_ConfigurationSettings.uptimeStatistics  ) Content += 'U';
    if ( m_ConfigurationSettings.errorLogs         ) Content += 'e';
    if ( m_ConfigurationSettings.environmentVars   ) Content += 'E';
    if ( m_ConfigurationSettings.regionalSettings  ) Content += 'R';
    if ( m_ConfigurationSettings.windowsNetwork    ) Content += 'N';
    if ( m_ConfigurationSettings.networkTcpIp      ) Content += 't';
    if ( m_ConfigurationSettings.hardwareDevices   ) Content += 'z';
    if ( m_ConfigurationSettings.displays          ) Content += 'D';
    if ( m_ConfigurationSettings.displayAdapters   ) Content += 'a';
    if ( m_ConfigurationSettings.installedPrinters ) Content += 'I';
    if ( m_ConfigurationSettings.biosVersion       ) Content += 'b';
    if ( m_ConfigurationSettings.systemManagement  ) Content += 'M';
    if ( m_ConfigurationSettings.processors        ) Content += 'p';
    if ( m_ConfigurationSettings.memory            ) Content += 'm';
    if ( m_ConfigurationSettings.physicalDisks     ) Content += 'i';
    if ( m_ConfigurationSettings.drives            ) Content += 'd';
    if ( m_ConfigurationSettings.commPorts         ) Content += 'c';
    if ( m_ConfigurationSettings.startupPrograms   ) Content += 'S';
    if ( m_ConfigurationSettings.services          ) Content += 'A';
    if ( m_ConfigurationSettings.runningPrograms   ) Content += 'r';
    if ( m_ConfigurationSettings.odbcInformation   ) Content += 'C';
    if ( m_ConfigurationSettings.oleDbProviders    ) Content += 'O';
    if ( m_ConfigurationSettings.softwareMetering  ) Content += 'H';
    if ( m_ConfigurationSettings.userLogonStats    ) Content += 'G';
    Content += PXS_STRING_CRLF;

    // Database
    Content += L"reportShowSQL=";
    if ( m_ConfigurationSettings.reportShowSQL )
    {
        Content += L"1";
    }
    else
    {
        Content += L"0";
    }
    Content += PXS_STRING_CRLF;

    Content += L"maxErrorRate=";
    Content += Format.UInt32( m_ConfigurationSettings.maxErrorRate );
    Content += PXS_STRING_CRLF;

    Content += L"maxAffectedRows=";
    Content += Format.UInt32( m_ConfigurationSettings.maxAffectedRows );
    Content += PXS_STRING_CRLF;

    Content += L"connectTimeoutSecs=";
    Content += Format.UInt32( m_ConfigurationSettings.connectTimeoutSecs );
    Content += PXS_STRING_CRLF;

    Content += L"queryTimeoutSecs=";
    Content += Format.UInt32( m_ConfigurationSettings.queryTimeoutSecs );
    Content += PXS_STRING_CRLF;

    Content += L"reportMaxRecords=";
    Content += Format.UInt32( m_ConfigurationSettings.reportMaxRecords );
    Content += PXS_STRING_CRLF;

    Content += L"DBMS=";
    Content += m_ConfigurationSettings.DBMS;
    Content += PXS_STRING_CRLF;

    Content += L"DatabaseName=";
    Content += m_ConfigurationSettings.DatabaseName;
    Content += PXS_STRING_CRLF;

    Content += L"MySqlDriver=";
    Content += m_ConfigurationSettings.MySqlDriver;
    Content += PXS_STRING_CRLF;

    Content += L"PostgreSqlDriver=";
    Content += m_ConfigurationSettings.PostgreSqlDriver;
    Content += PXS_STRING_CRLF;

    Content += L"LastReportName=";
    Content += m_ConfigurationSettings.LastReportName;
    Content += PXS_STRING_CRLF;

    // Other
    GetBounds( &bounds );
    Content += L"AppFrameBounds=";
    Content += Format.Int32( bounds.left );
    Content += L",";
    Content += Format.Int32( bounds.top );
    Content += L",";
    Content += Format.Int32( bounds.right );
    Content += L",";
    Content += Format.Int32( bounds.bottom );
    Content += PXS_STRING_CRLF;

    Content += PXS_STRING_CRLF;
    Content += L"!EOF";

    PXSGetExeDirectory( &IniPath );
    IniPath += PXS_WINAUDIT_INI;
    FileObject.CreateNew( IniPath, 0, true );
    FileObject.WriteChars( Content );
}
