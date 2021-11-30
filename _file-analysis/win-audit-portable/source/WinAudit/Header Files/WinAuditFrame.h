///////////////////////////////////////////////////////////////////////////////////////////////////
//
// WinAudit Frame Class Header
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

#ifndef WINAUDIT_WINAUDIT_FRAME_H_
#define WINAUDIT_WINAUDIT_FRAME_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "WinAudit/Header Files/WinAudit.h"

// 2. C System Files
#include <signal.h>
#include <time.h>

// 3. C++ System Files

// 4. Other Libraries
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/FindTextBar.h"
#include "PxsBase/Header Files/Frame.h"
#include "PxsBase/Header Files/ImageButton.h"
#include "PxsBase/Header Files/MenuBar.h"
#include "PxsBase/Header Files/MenuItem.h"
#include "PxsBase/Header Files/MenuPopup.h"
#include "PxsBase/Header Files/Mutex.h"
#include "PxsBase/Header Files/ProgressBar.h"
#include "PxsBase/Header Files/RichEditBox.h"
#include "PxsBase/Header Files/Splitter.h"
#include "PxsBase/Header Files/StatusBar.h"
#include "PxsBase/Header Files/TabWindow.h"
#include "PxsBase/Header Files/TArray.h"
#include "PxsBase/Header Files/ToolBar.h"
#include "PxsBase/Header Files/TreeView.h"
#include "PxsBase/Header Files/TreeViewItem.h"

// 5. This Project
#include "WinAudit/Header Files/AuditRecord.h"
#include "WinAudit/Header Files/AuditThread.h"
#include "WinAudit/Header Files/AuditThreadParameter.h"
#include "WinAudit/Header Files/ConfigurationSettings.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Application specific messages over and above any from PxsBase
///////////////////////////////////////////////////////////////////////////////////////////////////

// File menu
const WORD PXS_APP_MSG_AUDIT_START          = WM_APP + 100;
const WORD PXS_APP_MSG_SAVE                 = WM_APP + 101;
const WORD PXS_APP_MSG_SEND_EMAIL           = WM_APP + 102;
const WORD PXS_APP_MSG_ODBC_EXPORT          = WM_APP + 103;
const WORD PXS_APP_MSG_DESKTOP_SHORTCUT     = WM_APP + 104;
const WORD PXS_APP_MSG_EXIT                 = WM_APP + 105;

// Edit menu
const WORD PXS_APP_MSG_EDIT_FIND            = WM_APP + 110;
const WORD PXS_APP_MSG_EDIT_SET_FONT        = WM_APP + 111;
const WORD PXS_APP_MSG_EDIT_SMALLER_TEXT    = WM_APP + 112;
const WORD PXS_APP_MSG_EDIT_BIGGER_TEXT     = WM_APP + 113;

// View menu
const WORD PXS_APP_MSG_TOGGLE_TOOLBAR       = WM_APP + 120;
const WORD PXS_APP_MSG_TOGGLE_STATUSBAR     = WM_APP + 121;
const WORD PXS_APP_MSG_VIEW_OPTIONS         = WM_APP + 122;
const WORD PXS_APP_MSG_VIEW_DISK_INFO       = WM_APP + 123;
const WORD PXS_APP_MSG_VIEW_DISPLAY_INFO    = WM_APP + 124;
const WORD PXS_APP_MSG_VIEW_FIRMWARE_INFO   = WM_APP + 125;
const WORD PXS_APP_MSG_VIEW_PROCESSOR_INFO  = WM_APP + 126;
const WORD PXS_APP_MSG_VIEW_POLICY_INFO     = WM_APP + 127;
const WORD PXS_APP_MSG_VIEW_SOFTWARE_INFO   = WM_APP + 128;

// Language menu
const WORD PXS_APP_MSG_LANGUAGE_CHINESE_TW  = WM_APP + 140;
const WORD PXS_APP_MSG_LANGUAGE_CZECH       = WM_APP + 141;
const WORD PXS_APP_MSG_LANGUAGE_DANISH      = WM_APP + 142;
const WORD PXS_APP_MSG_LANGUAGE_GERMAN      = WM_APP + 143;
const WORD PXS_APP_MSG_LANGUAGE_ENGLISH     = WM_APP + 144;
const WORD PXS_APP_MSG_LANGUAGE_SPANISH     = WM_APP + 145;
const WORD PXS_APP_MSG_LANGUAGE_GREEK       = WM_APP + 146;
const WORD PXS_APP_MSG_LANGUAGE_FRENCH_BE   = WM_APP + 147;
const WORD PXS_APP_MSG_LANGUAGE_FRENCH_FR   = WM_APP + 148;
const WORD PXS_APP_MSG_LANGUAGE_HEBREW      = WM_APP + 149;
const WORD PXS_APP_MSG_LANGUAGE_INDONESIAN  = WM_APP + 150;
const WORD PXS_APP_MSG_LANGUAGE_ITALIAN     = WM_APP + 151;
const WORD PXS_APP_MSG_LANGUAGE_JAPANESE    = WM_APP + 152;
const WORD PXS_APP_MSG_LANGUAGE_KOREAN      = WM_APP + 153;
const WORD PXS_APP_MSG_LANGUAGE_HUNGARIAN   = WM_APP + 154;
const WORD PXS_APP_MSG_LANGUAGE_DUTCH       = WM_APP + 155;
const WORD PXS_APP_MSG_LANGUAGE_POLISH      = WM_APP + 156;
const WORD PXS_APP_MSG_LANGUAGE_PORTUGESE_BR= WM_APP + 157;
const WORD PXS_APP_MSG_LANGUAGE_PORTUGESE_PT= WM_APP + 158;
const WORD PXS_APP_MSG_LANGUAGE_RUSSIAN     = WM_APP + 159;
const WORD PXS_APP_MSG_LANGUAGE_SLOVAK      = WM_APP + 160;
const WORD PXS_APP_MSG_LANGUAGE_SERBIAN     = WM_APP + 161;
const WORD PXS_APP_MSG_LANGUAGE_FINNISH     = WM_APP + 162;
const WORD PXS_APP_MSG_LANGUAGE_THAI        = WM_APP + 163;
const WORD PXS_APP_MSG_LANGUAGE_TURKISH     = WM_APP + 164;

// Help menu
const WORD PXS_APP_MSG_HELP                 = WM_APP + 170;
const WORD PXS_APP_MSG_START_BROWSER        = WM_APP + 171;
const WORD PXS_APP_MSG_ABOUT_DIALOG         = WM_APP + 172;
const WORD PXS_APP_MSG_HELP_CREATE_GUID     = WM_APP + 173;
const WORD PXS_APP_MSG_TOGGLE_LOGFILE       = WM_APP + 174;

// Other
const WORD PXS_APP_MSG_AUDIT_COMMAND_LINE   = WM_APP + 180;
const WORD PXS_APP_MSG_AUDIT_THREAD_UPDATE  = WM_APP + 181;
const WORD PXS_APP_MSG_AUDIT_THREAD_DONE    = WM_APP + 182;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class WinAuditFrame : public Frame
{
    public:
        // Default constructor
        WinAuditFrame();

        // Destructor
        ~WinAuditFrame();

        // Methods
        void    AuditStart();
        DWORD   AuditThreadCallback( DWORD timeRemaining,
                                     DWORD percentDone,
                                     const TArray< AuditRecord >* pAuditRecords,
                                     const Exception* pAuditThreadException );
        void    AuditThreadDone();
        void    AuditThreadUpdate();
        void    CreateControls();
        void    CreateMenuItems();
        void    DoAuditInCommandLineMode();
        void    DoLayout();
        bool    GetAutoStartAudit();
        void    ReadIniFile();
        void    SetControlsAndMenus();
        void    SetUserInterfaceLanguage( WORD primaryLanguage, WORD subLanguage );

    protected:
        // Data members

        // Methods
        void    AppMessageEvent( UINT uMsg, WPARAM wParam, LPARAM lParam );
        void    CloseEvent();
        LRESULT CommandEvent( WPARAM wParam, LPARAM lParam );
        void    CopyDataEvent( HWND hWnd, const COPYDATASTRUCT* pCds );
        void    HelpEvent( const HELPINFO* pHelpInfo );
        void    NotifyEvent( const NMHDR* pNmhdr );
        void    SettingChangeEvent();

    private:
        // Copy constructor - not allowed
        WinAuditFrame( const WinAuditFrame& oWinAuditFrame );

        // Assignment operator - not allowed
        WinAuditFrame& operator= ( const WinAuditFrame& oWinAuditFrame );

        // Methods
        void    CopySelection();
        void    DesktopShortCut();
        void    DestroyMenuItems();
        void    FillHelpContents();
        void    FindText( bool forward, bool fromSelectionStart );
        void    HandleButtonClick( HWND hWnd );
        void    HandleHideWindow( HWND hWnd );
        void    HandleItemSelected( HWND hWnd );
        void    NavigateToWebSite();
        void    Save();
        void    SaveToFile( size_t tabIndex, DWORD dataFormat, const String& FilePath );
        void    SetApplicationBounds( const String& ApplicationBounds );
        void    SelectAllContents();
        void    SelectFont();
        void    SendEmail();
        void    SetBiggerSmallerText( bool bigger );
        void    SetControlLabels();
        void    SetEditMenu();
        void    SetMenuLabels();
        void    SetColours();
        void    SetReportOptions( const String& ReportOptions );
        void    ShowAboutDialog();
        void    ShowCommandLineHelpMessage();
        void    ShowCreateGuid();
        void    ShowHideHelp( bool show );
        void    ShowHideSidePanel( bool show );
        void    ShowOdbcExport();
        void    StartStopLogger( bool start );
        void    UpdateLogWindow();
        void    ViewDisplayInformation();
        void    ViewDiskInformation();
        void    ViewFirmwareInformation();
        void    ViewOptionsDialog();
        void    ViewPolicyInformation();
        void    ViewProcessorInformation();
        void    ViewSoftwareInformation();
        void    WriteIniFile();

        // Data members
        bool         m_bAuditing;
        bool         m_bCreatedReport;
        bool         m_bSavedReport;
        bool         m_bCreatedControls;
        DWORD        m_uSaveAuditFilterIndex;
        DWORD        m_uTableCounter;
        const size_t TAB_INDEX_AUDIT;
        const size_t TAB_INDEX_PROCESSORS;
        const size_t TAB_INDEX_DISKS;
        const size_t TAB_INDEX_DISPLAYS;
        const size_t TAB_INDEX_FIRMWARE;
        const size_t TAB_INDEX_SOFTWARE;
        const size_t TAB_INDEX_HELP;
        const size_t TAB_INDEX_LOGGER;
        const DWORD  DATA_FORMAT_CSV;
        const DWORD  DATA_FORMAT_CSV2;
        const DWORD  DATA_FORMAT_RTF;
        const DWORD  DATA_FORMAT_HTML;
        const DWORD  DATA_FORMAT_DEFAULT;
        AuditThread  m_AuditThread;
        ConfigurationSettings m_ConfigurationSettings;
        TArray< AuditRecord > m_AuditRecords;

        // Multi-thread variables
        sig_atomic_t          m_bRunAuditThreadMT;
        DWORD                 m_uAuditPercentDoneMT;
        AuditThreadParameter  m_AuditThreadParameterMT;
        TArray< Exception >   m_AuditThreadExceptionsMT;
        TArray< AuditRecord > m_NewAuditRecordsMT;
        Mutex                 m_Mutex;

        // Menu Items
        MenuBar         m_MenuBar;
            MenuItem        m_FileMenuBarItem;
                MenuPopup       m_FilePopup;
                MenuItem        m_FileAudit;
                MenuItem        m_FileSave;
                MenuItem        m_FileEmail;
                MenuItem        m_FileOdbc;
                MenuItem        m_FileShortcut;
                MenuItem        m_FileExit;

            MenuItem        m_EditMenuBarItem;
                MenuPopup       m_EditPopup;
                MenuItem        m_EditSelectAll;
                MenuItem        m_EditCopy;
                MenuItem        m_EditFind;
                MenuItem        m_EditSetFont;
                MenuItem        m_EditBiggerText;
                MenuItem        m_EditSmallerText;

            MenuItem        m_ViewMenuBarItem;
                MenuPopup       m_ViewPopup;
                MenuItem        m_ViewToolBar;
                MenuItem        m_ViewStatusBar;
                MenuItem        m_ViewOptions;
                MenuItem        m_ViewDiskInfo;
                MenuItem        m_ViewDisplayInfo;
                MenuItem        m_ViewFirmwareInfo;
                MenuItem        m_ViewPolicyInfo;
                MenuItem        m_ViewProcessorInfo;
                MenuItem        m_ViewSoftwareInfo;

            MenuItem        m_LanguageMenuBarItem;
                MenuPopup       m_LanguagePopup;
                MenuItem        m_LanguageChineseTw;
                MenuItem        m_LanguageCzech;
                MenuItem        m_LanguageDanish;
                MenuItem        m_LanguageGerman;
                MenuItem        m_LanguageEnglish;
                MenuItem        m_LanguageSpanish;
                MenuItem        m_LanguageGreek;
                MenuItem        m_LanguageBelgian;
                MenuItem        m_LanguageFrench;
                MenuItem        m_LanguageHebrew;
                MenuItem        m_LanguageItalian;
                MenuItem        m_LanguageIndonesian;
                MenuItem        m_LanguageJapanese;
                MenuItem        m_LanguageKorean;
                MenuItem        m_LanguageHungarian;
                MenuItem        m_LanguageDutch;
                MenuItem        m_LanguagePolish;
                MenuItem        m_LanguageBrazilian;
                MenuItem        m_LanguagePortuguese;
                MenuItem        m_LanguageRussian;
                MenuItem        m_LanguageSlovak;
                MenuItem        m_LanguageSerbian;
                MenuItem        m_LanguageFinnish;
                MenuItem        m_LanguageThai;
                MenuItem        m_LanguageTurkish;

            MenuItem        m_HelpMenuBarItem;
                MenuPopup       m_HelpPopup;
                MenuItem        m_HelpUsingWinAudit;
                MenuItem        m_HelpWebsite;
                MenuItem        m_HelpAbout;
                MenuItem        m_HelpCreateGuid;
                MenuItem        m_HelpLogFile;

        // Windows
        Window          m_MainPanel;
            ToolBar         m_ToolBar;
                ImageButton     m_AuditButton;
                ImageButton     m_StopButton;
                ImageButton     m_OptionsButton;
                ImageButton     m_SaveButton;
                ImageButton     m_EmailButton;
                ImageButton     m_HelpButton;
            Window          m_WorkSpace;
                FindTextBar     m_FindTextBar;
                Splitter        m_Splitter;
                Window          m_SidePanel;
                    TreeView        m_AuditCategories;
                    TreeView        m_HelpContents;
                TabWindow       m_TabWindow;
                    RichEditBox     m_AuditRichBox;
                    RichEditBox     m_DiskRichBox;
                    RichEditBox     m_DisplayRichBox;
                    RichEditBox     m_FirmwareRichBox;
                    RichEditBox     m_ProcessorRichBox;
                    RichEditBox     m_SoftwareRichBox;
                    RichEditBox     m_HelpRichBox;
                    RichEditBox     m_AboutRichBox;
                    RichEditBox     m_LogRichBox;
            StatusBar       m_StatusBar;
                ProgressBar     m_ProgressBar;
};

#endif  // WINAUDIT_WINAUDIT_FRAME_H_
