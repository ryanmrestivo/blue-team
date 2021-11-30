///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Splash Screen Class Implementation
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
#include "PxsBase/Header Files/SplashScreen.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Font.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/StaticControl.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
SplashScreen::SplashScreen()
             :TIMER_ID( 1002 )
{
    int  width    = 0, height = 0;
    RECT workArea = { 0, 0, 0, 0 };

    // Size
    if ( SystemParametersInfo( SPI_GETWORKAREA, 0, &workArea, 0 ) )
    {
        width  = workArea.right  - workArea.left;
        height = workArea.bottom - workArea.top;
    }
    if ( width  > 640 ) width  = 640;
    if ( height > 480 ) height = 480;

    // Creation parameters
    m_CreateStruct.cx         =  3 * (width  / 5);
    m_CreateStruct.cy         =  2 * (height / 4);
    m_CreateStruct.style      = static_cast<LONG>( WS_POPUP );  // = 0x80000000
    m_CreateStruct.dwExStyle |= WS_EX_TOOLWINDOW;

    // Properties
    try
    {
        SetBorderStyle( PXS_SHAPE_SHADOW );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy constructor - not allowed so no implementation

// Destructor
SplashScreen::~SplashScreen()
{
    if ( m_hWindow )
    {
        KillTimer( m_hWindow, TIMER_ID );
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
//      Show the splash screen
//
//  Parameters:
//      hWndOwner         - owner  window, if NULL will use the desktop
//      displayTime       - time to display splash screen in milliseconds
//      bitmapResourceID  - the resource id of the bitmap
//      alpha             - alpha transparency
//      OwnerLabel        - locale string for "Registered Owner:"
//      RegisteredOrganisationLabel - locale string for
//                                    "Registered Organisation:"
//      ProductIDLabel    - locale string for "Product ID:"
//
//  Returns:
//     void
//===============================================================================================//
void SplashScreen::Show( HWND hWndOwner, UINT displayTime, WORD bitmapResourceID, BYTE alpha )
{
    int     xPos = 0, yPos = 0, lineHeight = 25;
    Font    SplashFont;
    RECT    bounds = { 0, 0, 0, 0 }, workArea = { 0, 0, 0, 0 };
    SIZE    windowSize = { 0, 0 };
    HICON   hIcon    = nullptr;
    HBITMAP hBitmap  = nullptr;
	HMODULE hModule;
    String  Text, ExePath, OwnerLabel, OrganisationLabel, ProductIdLabel;
    String  CopyrightNotice, RegisteredOrgName, RegisteredOwnerName;
    String  ProductID, SupportEmail, WebSiteURL;
    StaticControl Static;

    // Set display time if it is not zero
    if ( displayTime == 0 )
    {
        return;
    }
    GetSize( &windowSize );
    if ( g_pApplication )
    {
        g_pApplication->GetManufacturerInfo( &CopyrightNotice,
                                             nullptr,
                                             nullptr,
                                             &RegisteredOrgName,
                                             &RegisteredOwnerName,
                                             &ProductID, &SupportEmail, &WebSiteURL );
    }


    // Icon - use the first icon in the exe file. Observed access violation in
    // ExtractIcon if path is null
    PXSGetExeDirectory( &ExePath );
    hModule = GetModuleHandle( nullptr );
    if ( ExePath.c_str() && hModule )
    {
        hIcon = ExtractIcon( static_cast<HINSTANCE>( hModule ), ExePath.c_str(), 0 );
        if ( hIcon > (HICON)1 )
        {
            SetRect( &bounds, 10, 10, 50, 45 );
            Static.SetBounds( bounds );
            Static.SetIcon( hIcon );
            Static.SetShape( PXS_SHAPE_SHADOW );
            Static.SetShapeColour( PXS_COLOUR_NAVY );
            Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
            Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
            m_Statics.Add( Static );
            Static.Reset();
            DestroyIcon( hIcon );
        }
    }

    // Title, centred in window
    Text = PXS_STRING_EMPTY;
    PXSGetApplicationName( &Text );
    SplashFont.SetPointSize( 12, nullptr );
    SplashFont.SetBold( true );
    SplashFont.SetUnderlined( true );
    SplashFont.Create();
    Static.SetFont( SplashFont );

    bounds.left   = 50;
    bounds.top    = 10;
    bounds.bottom = 50;
    bounds.right  = windowSize.cx;
    Static.SetBounds( bounds );
    Static.SetFont( SplashFont );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    Static.SetText( Text );
    m_Statics.Add( Static );
    Static.Reset();

    // Reset font
    SplashFont.SetPointSize( 9, nullptr );
    SplashFont.SetBold( false );
    SplashFont.SetUnderlined( false );
    SplashFont.Create();

    // Registered Owner
    PXSGetResourceString( PXS_IDS_510_REGISTERED_OWNER, &OwnerLabel );
    bounds.left   = 40;
    bounds.top    = bounds.bottom;
    bounds.right  = windowSize.cx / 2;
    bounds.bottom = bounds.top + lineHeight;
    Static.SetBounds( bounds );
    Static.SetFont( SplashFont );
    Static.SetAlignmentX( PXS_LEFT_ALIGNMENT );
    Static.SetText( OwnerLabel );
    m_Statics.Add( Static );
    Static.Reset();

    bounds.left   = bounds.right;
    bounds.right  = windowSize.cx - 40;
    Static.SetBounds( bounds );
    Static.SetFont( SplashFont );
    Static.SetAlignmentX( PXS_RIGHT_ALIGNMENT );
    Static.SetText( RegisteredOwnerName );
    Static.SetFont( SplashFont );
    m_Statics.Add( Static );
    Static.Reset();

    // Registered Organization
    PXSGetResourceString( PXS_IDS_511_REGISTERED_ORG, &OrganisationLabel );
    bounds.left   = 40;
    bounds.top    = bounds.bottom;
    bounds.right  = windowSize.cx / 2;
    bounds.bottom = bounds.top + lineHeight;
    Static.SetBounds( bounds );
    Static.SetFont( SplashFont );
    Static.SetAlignmentX( PXS_LEFT_ALIGNMENT );
    Static.SetText( OrganisationLabel );
    m_Statics.Add( Static );
    Static.Reset();

    bounds.left   = bounds.right;
    bounds.right  = windowSize.cx - 40;
    Static.SetBounds( bounds );
    Static.SetFont( SplashFont );
    Static.SetAlignmentX( PXS_RIGHT_ALIGNMENT );
    Static.SetText( RegisteredOrgName );
    m_Statics.Add( Static );
    Static.Reset();

    // Product ID
    PXSGetResourceString( PXS_IDS_512_PRODUCT_ID, &ProductIdLabel );
    bounds.left   = 40;
    bounds.top    = bounds.bottom;
    bounds.right  = windowSize.cx / 2;
    bounds.bottom = bounds.top + lineHeight;
    Static.SetBounds( bounds );
    Static.SetFont( SplashFont );
    Static.SetAlignmentX( PXS_LEFT_ALIGNMENT );
    Static.SetText( ProductIdLabel );
    m_Statics.Add( Static );
    Static.Reset();

    bounds.left   = bounds.right;
    bounds.right  = windowSize.cx - 40;
    Static.SetBounds( bounds );
    Static.SetFont( SplashFont );
    Static.SetAlignmentX( PXS_RIGHT_ALIGNMENT );
    Static.SetText( ProductID );
    m_Statics.Add( Static );
    Static.Reset();

    // Email
    bounds.top    = bounds.bottom;
    bounds.bottom = bounds.top + lineHeight;
    Static.SetBounds( bounds );
    Static.SetFont( SplashFont );
    Static.SetAlignmentX( PXS_RIGHT_ALIGNMENT );
    Static.SetText( SupportEmail );
    Static.SetHyperlink( true );
    m_Statics.Add( Static );
    Static.Reset();

    // Website
    bounds.top    = bounds.bottom;
    bounds.bottom = bounds.top + lineHeight;
    Static.SetBounds( bounds );
    Static.SetFont( SplashFont );
    Static.SetAlignmentX( PXS_RIGHT_ALIGNMENT );
    Static.SetText( WebSiteURL );
    m_Statics.Add( Static );
    Static.Reset();

    // Horizontal line
    bounds.left   = 10;
    bounds.right  = windowSize.cx - 10;
    bounds.top    = windowSize.cy - lineHeight - 5;
    bounds.bottom = bounds.top + 1;
    Static.SetBounds( bounds );
    Static.SetFont( SplashFont );
    Static.SetShape( PXS_SHAPE_RECTANGLE );
    Static.SetShapeColour( PXS_COLOUR_NAVY );
    m_Statics.Add( Static );
    Static.Reset();

    // Set small font size
    SplashFont.SetPointSize( 8, nullptr );

    // Copyright Notice
    Text = PXS_STRING_EMPTY;
    bounds.top    = bounds.bottom + 5;
    bounds.bottom = bounds.top + lineHeight;
    Static.SetBounds( bounds );
    Static.SetFont( SplashFont );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    Static.SetForeground( PXS_COLOUR_GREY );
    Static.SetText( CopyrightNotice );
    m_Statics.Add( Static );
    Static.Reset();

    // Image on bottom right
    if ( bitmapResourceID )
    {
        hBitmap = LoadBitmap( GetModuleHandle( nullptr ),  MAKEINTRESOURCE( bitmapResourceID ) );
        if ( hBitmap )
        {
            try
            {
                // Set the position at the bottom right hand corner
                bounds.left   = 0;
                bounds.top    = bounds.bottom;
                bounds.right  = windowSize.cx;
                bounds.bottom = windowSize.cy;
                Static.SetBounds( bounds );
                Static.SetAlignmentX( PXS_RIGHT_ALIGNMENT );
                Static.SetAlignmentY( PXS_BOTTOM_ALIGNMENT );
                Static.SetBitmap( hBitmap );
                m_Statics.Add( Static );
                Static.Reset();
            }
            catch ( const Exception& eBitmap )
            {
               PXSLogException( eBitmap, __FUNCTION__ );
            }
        }
    }

    // Create the window, if no owner use the desktop
    if ( hWndOwner == nullptr )
    {
        hWndOwner = GetDesktopWindow();
    }
    SetVisible( false );    // Want it to be initially invisible
    Create( hWndOwner );
    if ( m_hWindow )
    {
        SetLayeredWindowAttributes( m_hWindow,
                                    PXS_COLOUR_WHITE, alpha, LWA_ALPHA );
        // Centre it on the desktop
        SystemParametersInfo(SPI_GETWORKAREA, 0, &workArea, 0);
        xPos = (workArea.right - workArea.left - windowSize.cx) / 2;
        yPos = (workArea.bottom - workArea.top - windowSize.cy) / 2;
        xPos = PXSMaxInt(xPos, 0);
        yPos = PXSMaxInt(yPos, 0);
        MoveWindow(m_hWindow, xPos, yPos, windowSize.cx, windowSize.cy, FALSE);
        SetVisible(true);
        Repaint();
        SetTimer( m_hWindow, TIMER_ID, displayTime, nullptr );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Handle WM_TIMER event.
//
//  Parameters:
//      timerID - The identifier of the timer that fired this event
//
//  Returns:
//      void.
//===============================================================================================//
void SplashScreen::TimerEvent( UINT_PTR timerID )
{
    // Hide splash screen
    if ( timerID == TIMER_ID )
    {
        if ( m_hWindow )
        {
            KillTimer( m_hWindow, TIMER_ID );
        }
        SetVisible( false );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
