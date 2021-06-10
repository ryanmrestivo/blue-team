///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Unhandled Exception Dialog Class Implementation
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
#include "PxsBase/Header Files/UnhandledExceptionDialog.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Font.h"
#include "PxsBase/Header Files/Mail.h"
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/StaticControl.h"
#include "PxsBase/Header Files/SystemInformation.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
UnhandledExceptionDialog::UnhandledExceptionDialog()
                         :m_MailButton(),
                          m_ExitButton(),
                          m_DetailsText(),
                          m_ExceptionMessage()

{
    try
    {
        String ApplicationName;

        SetSize( 550, 365 );
        PXSGetApplicationName( &ApplicationName );
        SetTitle( ApplicationName );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy constructor - not allowed so no implementation

// Destructor
UnhandledExceptionDialog::~UnhandledExceptionDialog()
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
//      Set the string to show in the details box
//
//  Parameters:
//      ExceptionMessage - pointer to string
//
//  Remarks:
//      Does not throw
//
//  Returns:
//      void
//===============================================================================================//
void UnhandledExceptionDialog::SetExceptionMessage( const String& ExceptionMessage )
{
    try
    {
        m_ExceptionMessage = ExceptionMessage;
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Handle WM_COMMAND events.
//
//  Parameters:
//        wParam -
//        lParam -
//
//  Returns:
//        0 if handled, else non-zero.
//===============================================================================================//
LRESULT UnhandledExceptionDialog::CommandEvent( WPARAM wParam, LPARAM lParam )
{
    WORD id = LOWORD(wParam);  // Item, control, or accelerator identifier

    try
    {
        if ( ( id == IDCANCEL ) ||
             IsClickFromButton( m_ExitButton, wParam, lParam ) )
        {
            EndDialog( m_hWindow, IDCANCEL );
        }
        else if ( IsClickFromButton( m_MailButton, wParam, lParam ) )
        {
            Mail   MAPI;
            String Subject, SupportEmail, NoteText, AttachFilePath;

            // Send mail, catch exceptions as need to end the dialog
            try
            {
                if ( g_pApplication )
                {
                    g_pApplication->GetSupportEmail( &SupportEmail );
                }
                m_DetailsText.GetText( &NoteText );
                Subject = L"Program Error Report";
                MAPI.SendSimpleMail( SupportEmail, Subject, NoteText, AttachFilePath );
            }
            catch ( const Exception& eMAPI )
            {
                MessageBox( m_hWindow, eMAPI.GetMessage().c_str(), nullptr, MB_OK | MB_ICONERROR );
            }
            EndDialog( m_hWindow, IDOK );
        }
    }
    catch ( const Exception& e )
    {
        MessageBox( m_hWindow, e.GetMessage().c_str(), nullptr, MB_OK | MB_ICONERROR );
        PXSLogException( e, __FUNCTION__ );
    }

    return 0;   // = Handled
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
void UnhandledExceptionDialog::InitDialogEvent()
{
    SIZE   buttonSize = { 0, 0 }, clientSize = { 0, 0 };
    RECT   bounds     = { 0, 0, 0, 0 };
    POINT  location   = { 0, 0 };
    Font   StaticFont, FixedWidthFont;
    String Text, Caption, ManufacturerAddress, Diagnostics;
    String SupportEmail, WebSiteURL;
    StaticControl Static;
    SystemInformation SystemInfo;

    // Only allow initialization once
    if ( m_bControlsCreated )
    {
        return;
    }
    m_bControlsCreated = true;

    if ( m_hWindow == nullptr )
    {
        return;
    }
    GetClientSize( &clientSize );
    FixedWidthFont.SetStockFont( ANSI_FIXED_FONT );

    // Mail Button, add first so it has focus
    m_MailButton.GetSize( &buttonSize );
    location.x = clientSize.cx - 30 - ( 2 * buttonSize.cx );
    location.y = clientSize.cy - 10 - buttonSize.cy;
    m_MailButton.SetLocation( location );
    m_MailButton.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_132_EMAIL, &Caption );
    m_MailButton.SetText( Caption );

    // Close button
    location.x += ( buttonSize.cx + 20 );
    m_ExitButton.SetLocation( location );
    m_ExitButton.Create( m_hWindow );
    PXSGetResourceString( PXS_IDS_133_EXIT, &Caption );
    m_ExitButton.SetText( Caption );

    // Horizontal line
    bounds.left   = 10;
    bounds.right  = clientSize.cx - 10;
    bounds.top    = location.y - 10;
    bounds.bottom = bounds.top + 2;
    Static.SetBounds( bounds );
    Static.SetShape( PXS_SHAPE_SUNK );
    m_Statics.Add( Static );
    Static.Reset();

    // Icon - use the error icon in the exe file
    // No need to clean up after LoadIcon() as its a shared resource
    HICON hIcon = LoadIcon( nullptr, IDI_ERROR );
    if ( hIcon > (HICON)1 )
    {
        bounds.left   = 10;
        bounds.right  = 50;
        bounds.top    = 5;
        bounds.bottom = 45;
        Static.SetBounds( bounds );
        Static.SetIcon( hIcon );
        Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
        Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
        m_Statics.Add( Static );
        Static.Reset();
    }

    // Title Caption, centred
    PXSGetResourceString( PXS_IDS_520_APPLICATION_ERROR, &Caption );
    bounds.left    = 0;
    bounds.right   = clientSize.cx;
    bounds.top     = 5;
    bounds.bottom  = 45;
    StaticFont.SetPointSize( 12, nullptr );
    StaticFont.SetBold( true );
    StaticFont.SetUnderlined( true );
    StaticFont.Create();
    Static.SetFont( StaticFont );
    Static.SetBounds( bounds );
    Static.SetText( Caption );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    m_Statics.Add( Static );
    Static.Reset();

    // Detected error string
    PXSGetResourceString( PXS_IDS_521_ERROR_WILL_EXIT, &Caption );
    bounds.left   = 0;
    bounds.right  = clientSize.cx;
    bounds.top    = 50;
    bounds.bottom = 85;
    StaticFont.SetPointSize( 9, nullptr );
    StaticFont.SetBold( false );
    StaticFont.SetUnderlined( false );
    StaticFont.Create();
    Static.SetFont( StaticFont );
    Static.SetBounds( bounds );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    Static.SetAlignmentY( PXS_TOP_ALIGNMENT );
    Static.SetText( Caption );
    m_Statics.Add( Static );
    Static.Reset();

    // Text box
    bounds.left   = 10;
    bounds.right  = clientSize.cx - 10;
    bounds.top    = 85;
    bounds.bottom = 250;
    m_DetailsText.SetBounds( bounds );
    m_DetailsText.SetStyle( ES_AUTOVSCROLL | WS_HSCROLL | WS_VSCROLL, true );
    m_DetailsText.Create( m_hWindow );
    m_DetailsText.SetFont( FixedWidthFont );  // Must set after creation

    // Details
    Text  = L"Error Information\r\n";
    Text += L"=================\r\n\r\n";

    // Add the exception message
    Text += m_ExceptionMessage;
    Text += PXS_STRING_CRLF;

    // Get formatted operating system/environment diagnostics
    try
    {
        if ( g_pApplication )
        {
            g_pApplication->GetDiagnostics( &Diagnostics );
        }
        Text += Diagnostics;
        Text += PXS_STRING_CRLF;
    }
    catch ( const Exception& )
    { }     // Ignore
    Text += PXS_STRING_CRLF;

    PXSGetResourceString( PXS_IDS_522_PRIVACY, &Caption );
    Text += Caption;
    Text += PXS_STRING_CRLF;
    Text += L"-------\r\n\r\n";
    PXSGetResourceString( PXS_IDS_523_INFORMATION_HEREIN, &Caption);
    Text += Caption;
    Text += PXS_STRING_CRLF;
    Text += PXS_STRING_CRLF;

    // Add contact details
    if ( g_pApplication )
    {
        g_pApplication->GetManufacturerInfo( nullptr,
                                             &ManufacturerAddress,
                                             nullptr,
                                             nullptr,
                                             nullptr,
                                             nullptr,
                                             &SupportEmail,
                                             &WebSiteURL );
    }
    Text += L"==========================================\r\n\r\n";
    Text += ManufacturerAddress;
    Text += PXS_STRING_CRLF;
    Text += SupportEmail;
    Text += PXS_STRING_CRLF;
    Text += WebSiteURL;
    Text += PXS_STRING_CRLF;
    Text += PXS_STRING_CRLF;
    Text += PXS_STRING_CRLF;

    // Set edit box's text
    m_DetailsText.SetText( Text );

    // Quality feed back message
    PXSGetResourceString(PXS_IDS_524_EMAIL_ERROR_MESSAGE, &Caption);
    bounds.left    = 10;
    bounds.right   = clientSize.cx - 10;
    bounds.top     = 255;
    bounds.bottom  = bounds.top + 65;
    StaticFont.SetPointSize( 8, nullptr );
    Static.SetBounds( bounds );
    Static.SetFont( StaticFont );
    Static.SetForeground( PXS_COLOUR_NAVY );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    Static.SetText( Caption );
    m_Statics.Add( Static );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
