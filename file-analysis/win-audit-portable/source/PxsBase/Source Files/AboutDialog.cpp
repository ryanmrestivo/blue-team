///////////////////////////////////////////////////////////////////////////////////////////////////
//
// About Dialog Class Implementation
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
#include "PxsBase/Header Files/AboutDialog.h"

// 2. C System Files
#include <stdint.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/FileVersion.h"
#include "PxsBase/Header Files/Font.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/StaticControl.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default Constructor
AboutDialog::AboutDialog()
            :m_Button(),
             m_TextArea()
{
    try
    {
        SetSize( 450, 375 );
        SetBackground( PXS_COLOUR_WHITE );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy constructor - not allowed so no implementation

// Destructor
AboutDialog::~AboutDialog()
{
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator - not allowed so no implementation

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////


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
LRESULT AboutDialog::CommandEvent( WPARAM wParam, LPARAM lParam )
{
    LRESULT result = 0;

    // Ensure controls have been created
    if ( m_bControlsCreated == false )
    {
        return 0;
    }

    if ( ( LOWORD( wParam ) == IDCANCEL ) ||
         ( IsClickFromButton( m_Button, wParam, lParam ) ) )
    {
        EndDialog( m_hWindow, IDCANCEL );
    }
    else
    {
        result = 1;     // Not handled
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
void AboutDialog::InitDialogEvent()
{
    const  int GAP = 10;
    bool   rtlReading;
    int    deltaX  = 0, staticWidth = 0, xPos = 0;
    Font   TitleFont;
    File   FileObject;
    BYTE*  pbExe      = nullptr;
    SIZE   buttonSize = { 0, 0 }, clientSize = { 0, 0 };
    RECT   bounds     = { 0, 0, 0, 0 };
    POINT  location   = { 0, 0 };
    DWORD  exeLength  = 0, alignmentX = 0;
    UINT64 fileSize   = 0;
    HICON  hIcon      = nullptr;
    String Manufacturer, ExePath, DataString, Text, ApplicationName;
    String Version, Description, Md5Hash, Caption, ProductID, TempString;
    String CopyrightNotice, RegisteredOrgName, RegisteredOwnerName, WebSiteURL;
    String ManufacturerAddress;
    Formatter     Format;
    FileVersion   FileVer;
    AllocateBytes ExeBytes;
    StaticControl Static;

    // Only allow initialization once
    if ( ( m_hWindow == nullptr ) || m_bControlsCreated )
    {
        return;
    }
    GetClientSize( &clientSize );

    rtlReading = IsRightToLeftReading();
    if ( g_pApplication )
    {
        g_pApplication->GetManufacturerInfo( &CopyrightNotice,
                                             &ManufacturerAddress,
                                             nullptr,
                                             &RegisteredOrgName,
                                             &RegisteredOwnerName,
                                             &ProductID, nullptr, &WebSiteURL );
    }

    // Close button - add first so it has focus
    PXSGetResourceString( PXS_IDS_130_CLOSE, &Caption );
    m_Button.GetSize( &buttonSize );
    location.x = clientSize.cx - GAP - buttonSize.cx;
    location.y = clientSize.cy - GAP - buttonSize.cy;
    m_Button.SetLocation( location );
    m_Button.Create( m_hWindow );
    m_Button.SetText( Caption );

    // Horizontal line
    Caption = PXS_STRING_EMPTY;
    bounds.left   = GAP;
    bounds.right  = clientSize.cx - GAP;
    bounds.top    = location.y - GAP;
    bounds.bottom = bounds.top + 2;
    Static.SetBounds( bounds );
    Static.SetText( Caption );
    Static.SetShape( PXS_SHAPE_SUNK );
    m_Statics.Add( Static );
    Static.Reset();

    // Icon - use the first icon in the exe file,
    hIcon = LoadIcon( GetModuleHandle( nullptr ), MAKEINTRESOURCE( 1 ) );
    if ( hIcon )
    {
        Static.SetIcon( hIcon );
        SetRect( &bounds, GAP, 5, 40 + GAP, 45 );
        Static.SetBounds( bounds );
        Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
        Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
        m_Statics.Add( Static );
        Static.Reset();
    }

    // Title Caption, centred
    PXSGetApplicationName( &Text );
    TitleFont.SetPointSize( 12, nullptr );
    TitleFont.SetBold( true );
    TitleFont.SetUnderlined( true );
    TitleFont.Create();
    SetRect( &bounds, 0, 5, clientSize.cx, 30 );
    Static.SetText( Text );
    Static.SetFont( TitleFont );
    Static.SetBounds( bounds );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    m_Statics.Add( Static );
    Static.Reset();

    // Evaluate the size of the statics and where starting from
    staticWidth = ( clientSize.cx - ( 3 * GAP ) ) / 2;
    xPos        = GAP;
    deltaX      = staticWidth + GAP;      // Next column is to the right
    alignmentX  = PXS_LEFT_ALIGNMENT;

    // Registered Owner
    PXSGetResourceString( PXS_IDS_510_REGISTERED_OWNER, &Caption );
    bounds.left   = xPos;
    bounds.right  = bounds.left + staticWidth;
    bounds.top    = bounds.bottom + GAP + GAP;
    bounds.bottom = bounds.top + 20;
    Static.SetBounds( bounds );
    Static.SetText( Caption );
    Static.SetAlignmentX( alignmentX );
    m_Statics.Add( Static );
    Static.Reset();

    bounds.left += deltaX;
    bounds.right = bounds.left + staticWidth;
    Static.SetBounds( bounds );
    Static.SetText( RegisteredOwnerName );
    Static.SetAlignmentX( alignmentX );
    m_Statics.Add( Static );

    // Registered Organization
    PXSGetResourceString( PXS_IDS_511_REGISTERED_ORG, &Caption );
    bounds.left   = xPos;
    bounds.right  = bounds.left + staticWidth;
    bounds.top    = bounds.bottom;
    bounds.bottom = bounds.top + 20;
    Static.SetBounds( bounds );
    Static.SetText( Caption );
    Static.SetAlignmentX( alignmentX );
    m_Statics.Add( Static );

    bounds.left   += deltaX;
    bounds.right  = bounds.left + staticWidth;
    Static.SetBounds( bounds );
    Static.SetText( RegisteredOrgName );
    Static.SetAlignmentX( alignmentX );
    m_Statics.Add( Static );

    // Product ID
    PXSGetResourceString( PXS_IDS_512_PRODUCT_ID, &Caption );
    bounds.left   = xPos;
    bounds.right  = bounds.left + staticWidth;
    bounds.top    = bounds.bottom;
    bounds.bottom = bounds.top + 20;
    Static.SetBounds( bounds );
    Static.SetText( Caption );
    Static.SetAlignmentX( alignmentX );
    m_Statics.Add( Static );

    bounds.left += deltaX;
    bounds.right = bounds.left + staticWidth;
    Static.SetBounds( bounds );
    Static.SetText( ProductID );
    Static.SetAlignmentX( alignmentX );
    m_Statics.Add( Static );

    // Create the text box
    bounds.left   = GAP;
    bounds.right  = clientSize.cx - GAP;
    bounds.top    = bounds.bottom + GAP;
    bounds.bottom = bounds.top + 120;
    m_TextArea.SetBounds( bounds );
    m_TextArea.SetReadOnly( true );
    m_TextArea.SetStyle( WS_HSCROLL | WS_VSCROLL, true );
    m_TextArea.Create( m_hWindow );

    // Add application name
    PXSGetResourceString( PXS_IDS_104_APPLICATION, &Caption );
    PXSGetApplicationName( &ApplicationName );
    Text = Caption;
    Text += PXS_CHAR_COLON;
    TempString.FixedWidth( 14, PXS_CHAR_SPACE );
    Text += PXS_CHAR_SPACE;
    Text += PXS_CHAR_TAB;
    Text += ApplicationName;
    Text += PXS_STRING_CRLF;

    // Version Number
    PXSGetExePath( &ExePath );
    PXSGetResourceString( PXS_IDS_513_VERSION, &Caption );
    TempString = Caption;
    TempString += PXS_CHAR_COLON;
    TempString.FixedWidth( 14, PXS_CHAR_SPACE );
    TempString += PXS_CHAR_SPACE;
    TempString += PXS_CHAR_TAB;
    Text += TempString;
    FileVer.GetVersion( ExePath, &Manufacturer, &Version, &Description );
    Text += Version;
    Text += PXS_STRING_CRLF;

    // File size
    PXSGetResourceString( PXS_IDS_514_SIZE, &Caption );
    FileObject.Open( ExePath, GENERIC_READ, FILE_SHARE_READ | FILE_SHARE_WRITE, 1, false );
    fileSize    = FileObject.GetSize();
    DataString  = Format.UInt64( fileSize );
    TempString  = Caption;
    TempString += PXS_CHAR_COLON;
    TempString.FixedWidth( 14, PXS_CHAR_SPACE );
    TempString += PXS_CHAR_SPACE;
    TempString += PXS_CHAR_TAB;
    Text       += TempString;
    if ( DataString.GetLength() )
    {
        PXSGetResourceString( PXS_IDS_134_BYTES, &Caption );
        Text += DataString;
        Text += PXS_CHAR_SPACE;
        Text += Caption;
    }
    Text += PXS_STRING_CRLF;

    // File creation time
    PXSGetResourceString( PXS_IDS_515_TIMESTAMP, &Caption );
    FileObject.GetTimes( ExePath, true, nullptr, nullptr, &DataString );
    TempString = Caption;
    TempString += PXS_CHAR_COLON;
    TempString += PXS_CHAR_SPACE;
    TempString += PXS_CHAR_TAB;
    Text += TempString;
    Text += DataString;
    Text += PXS_STRING_CRLF;

    // MD5
    TempString  = L"MD5";
    TempString += PXS_CHAR_COLON;
    TempString.FixedWidth( 14, PXS_CHAR_SPACE );
    TempString += PXS_CHAR_SPACE;
    TempString += PXS_CHAR_TAB;
    Text    += TempString;
    if ( fileSize < UINT32_MAX )    // Test before cast to narrower
    {
        exeLength = PXSCastUInt64ToUInt32( fileSize );
        pbExe = ExeBytes.New( exeLength );
        FileObject.Read( pbExe, exeLength );
        PXSGetMd5Hash( pbExe, exeLength, &Md5Hash );
        Text += Md5Hash;
    }
    Text += PXS_STRING_CRLF;
    FileObject.Close();

    // Add exe path
    PXSGetResourceString( PXS_IDS_120_LOCATION, &Caption );
    TempString = Caption;
    TempString += PXS_CHAR_COLON;
    TempString.FixedWidth( 14, PXS_CHAR_SPACE );
    TempString += PXS_CHAR_SPACE;
    TempString += PXS_CHAR_TAB;
    Text += TempString;
    Text += ExePath;
    Text += PXS_STRING_CRLF;
    m_TextArea.SetText( Text );

    // Legal notice
    PXSGetResourceString( PXS_IDS_516_COPYRIGHT_STRING, &Caption );
    Text  = Caption;
    Text += PXS_STRING_CRLF;
    Text += Caption;

    bounds.left   = GAP;
    bounds.right  = clientSize.cx - GAP;
    bounds.top    = bounds.bottom + GAP;
    bounds.bottom = bounds.top + 35;
    Static.SetBounds( bounds );
    Static.SetText( Caption );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    m_Statics.Add( Static );

    // Manufacturer Address
    bounds.top    = bounds.bottom;
    bounds.bottom = bounds.top + 20;
    Static.SetBounds( bounds );
    Static.SetText( ManufacturerAddress );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    m_Statics.Add( Static );

    // Copyright string
    bounds.top    = bounds.bottom;
    bounds.bottom = bounds.top + 20;
    Static.SetBounds( bounds );
    Static.SetText( CopyrightNotice );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    m_Statics.Add( Static );
    Static.Reset();

    // Mirror for RTL
    if ( rtlReading )
    {
        RtlStatics();
        m_Button.RtlMirror( clientSize.cx );
        m_TextArea.RtlMirror( clientSize.cx );
    }
    m_bControlsCreated = true;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
