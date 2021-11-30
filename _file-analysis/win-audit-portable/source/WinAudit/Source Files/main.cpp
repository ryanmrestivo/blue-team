///////////////////////////////////////////////////////////////////////////////////////////////////
//
// WinAudit Main Implementation
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
#include "WinAudit/Header Files/WinAudit.h"

// 2. C System Files

// 3. C++ System Files
#include <exception>

// 4. Other Libraries
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/SplashScreen.h"

// 5. This Project
#include "WinAudit/Header Files/Resources.h"
#include "WinAudit/Header Files/WinAuditFrame.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Main Entry Point
///////////////////////////////////////////////////////////////////////////////////////////////////
int APIENTRY WinMain( HINSTANCE /* hInstance */,
                      HINSTANCE /* hPrevInstance */,
                      LPSTR     lpCmdLine,
                      int       /* nCmdShow */ )
{
    bool    guiMode   = false;
    DWORD   errorCode = 0;
    wchar_t szErrorMsg[ 64 ] = { 0 };

    try
    {
        HWND   hWndFrame = nullptr;
        String CommandLine, OwnerLabel, OrganisationLabel, ProductIdLabel;
        SplashScreen  Splash;
        WinAuditFrame WinAudit;

        // Set termination handlers
        set_terminate( PXSTerminateHandler );
        if ( lpCmdLine && *lpCmdLine )
        {
            CommandLine = GetCommandLine();
            SetUnhandledExceptionFilter( PXSWriteUnhandledExceptionToLog );
        }
        else
        {
            guiMode = true;
            SetUnhandledExceptionFilter( PXSShowUnhandledExceptionDialog );
        }

        // Global application object
        g_pApplication = new Application;
        if ( g_pApplication == nullptr )
        {
            MessageBox( nullptr,
                        L"ERROR_NOT_ENOUGH_MEMORY", PXS_APPLICATION_NAME, MB_OK | MB_ICONSTOP );
            return ERROR_NOT_ENOUGH_MEMORY;
        }

        // Set the application's strings
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_CS, LANG_CZECH     , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_DA, LANG_DANISH    , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_DE, LANG_GERMAN    , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_EL, LANG_GREEK     , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_ES, LANG_SPANISH   , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_EN, LANG_ENGLISH   , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_FI, LANG_FINNISH   , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_FR_BE,
                                                  LANG_FRENCH, SUBLANG_FRENCH_BELGIAN );
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_FR_FR,
                                                  LANG_FRENCH, SUBLANG_FRENCH );
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_HE, LANG_HEBREW    , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_HU, LANG_HUNGARIAN , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_ID, LANG_INDONESIAN, 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_IT, LANG_ITALIAN   , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_JP, LANG_JAPANESE  , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_KO, LANG_KOREAN    , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_NL, LANG_DUTCH     , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_PL, LANG_POLISH    , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_PT_BR,
                                                  LANG_PORTUGUESE, SUBLANG_PORTUGUESE_BRAZILIAN );
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_PT_PT,
                                                  LANG_PORTUGUESE, SUBLANG_PORTUGUESE );
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_RU, LANG_RUSSIAN   , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_SR, LANG_SERBIAN   , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_SK, LANG_SLOVAK    , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_TH, LANG_THAI      , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_TR, LANG_TURKISH   , 0);
        g_pApplication->AddAppResourceLanguageID( IDR_STRINGS_WINAUDIT_ZH_TW,
                                                  LANG_CHINESE, SUBLANG_CHINESE_TRADITIONAL );
        g_pApplication->LoadStringFile( LANG_NEUTRAL, SUBLANG_DEFAULT );

        // Properties
        g_pApplication->SetApplicationName( PXS_APPLICATION_NAME );
        g_pApplication->SetManufacturerName( L"Parmavex Services" );
        g_pApplication->SetCopyrightNotice( PXS_COPYRIGHT_NOTICE );
        g_pApplication->SetManufacturerAddress( PXS_PARMAVEX_ADDRESS );
        g_pApplication->SetSupportEmail( L"winaudit@parmavex.co.uk" );
        g_pApplication->SetWebSiteURL( L"http://www.parmavex.co.uk/winaudit.html" );
        g_pApplication->SetRegisteredOrgName( L"Parmavex Services" );
        g_pApplication->SetRegisteredOwnerName( L"Parmavex Services" );
        g_pApplication->SetProductID( L"Freeware" );

        // COM
        PXSInitializeComOnThread();
        PXSCoInitializeSecurity();

        // GUI vs. command line mode
        if ( guiMode )
        {
            // Main frame
            WinAudit.Create( HWND_DESKTOP );    // HWND_DESKTOP = 0
            hWndFrame = WinAudit.GetHwnd();
            g_pApplication->SetHwndMainFrame( hWndFrame );

            // Show the main frame
            WinAudit.SetTitle( PXS_APPLICATION_NAME );
            WinAudit.SetIconImage( PXS_IDI_WINAUDIT_FREEWARE, true );
            WinAudit.CreateControls();
            WinAudit.CreateMenuItems();
            WinAudit.SetUserInterfaceLanguage( LANG_NEUTRAL, SUBLANG_DEFAULT );
            WinAudit.SetControlsAndMenus();

            // Show the splash screen
            if ( PXSIsHighColour( nullptr ) )
            {
                Splash.SetBackgroundGradient( RGB( 232, 232, 232 ), PXS_COLOUR_WHITE, true );
            }
            Splash.SetDoubleBuffered( true );
            Splash.Show( hWndFrame, 1000, 0, 0 );

            // Don't fail app startup on config problem
            try
            {
                WinAudit.ReadIniFile();
            }
            catch ( const Exception& eIni )
            {
                MessageBox( nullptr,
                            eIni.GetMessage().c_str(), PXS_APPLICATION_NAME, MB_ICONEXCLAMATION );
            };
            WinAudit.DoLayout();
            WinAudit.SetVisible( true );
            if ( WinAudit.GetAutoStartAudit() )
            {
                PostMessage( hWndFrame, WM_COMMAND, PXS_APP_MSG_AUDIT_START, 0 );
            }
            errorCode = WinAudit.StartMessageLoop();
        }
        else
        {
           WinAudit.DoAuditInCommandLineMode();
        }
    }
    catch ( const Exception& e )
    {
        errorCode = e.GetErrorCode();
        PXSLogException( e, __FUNCTION__ );
        if ( guiMode )
        {
            MessageBox( nullptr,
                        e.GetMessage().c_str(), PXS_APPLICATION_NAME, MB_ICONEXCLAMATION );
        }
    }

    CoUninitialize();
    if ( errorCode )
    {
        StringCchPrintf( szErrorMsg,
                         ARRAYSIZE( szErrorMsg ),
                         L"Application exiting with error %lu.", errorCode );
        PXSLogAppError( szErrorMsg );
    }
    delete g_pApplication;

    return static_cast<int>( errorCode );
}
