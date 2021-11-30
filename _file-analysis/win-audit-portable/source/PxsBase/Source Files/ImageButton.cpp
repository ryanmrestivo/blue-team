///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Image Button Class Implementation
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
#include "PxsBase/Header Files/ImageButton.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/ToolTip.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
ImageButton::ImageButton( int buttonWidth, int buttonHeight, int imageWidth, int imageHeight )
             :m_bMouseInside( false ),
              m_bMouseWasPressed( false ),
              m_bDisplayCaption( true ),
              m_uTimerID( 1001 ),
              m_ImageSize(),        // Non-op
              m_hIconOn( nullptr ),
              m_hIconOff( nullptr ),
              m_hIconDefault( nullptr ),
              m_hBitmapOn( nullptr ),
              m_hBitmapOff( nullptr ),
              m_hBitmapDefault( nullptr ),
              m_StaticImage(),
              m_StaticText()
{
    // CREATESTRUCT
    m_CreateStruct.cy = buttonHeight;
    m_CreateStruct.cx = buttonWidth;

    // This
    m_ImageSize.cx    = imageWidth;
    m_ImageSize.cy    = imageHeight;

    // Properties
    m_crBackground    = GetSysColor( COLOR_BTNFACE );
}

// Copy constructor - not allowed so no implementation

// Destructor
ImageButton::~ImageButton()
{
    // No need to destroy icons because they were loaded
    // using LoadIcon

    // Delete bitmaps
    if ( m_hBitmapOn      ) DeleteObject( m_hBitmapOn );
    if ( m_hBitmapOff     ) DeleteObject( m_hBitmapOff );
    if ( m_hBitmapDefault ) DeleteObject( m_hBitmapDefault );

    // Ensure the timer has been killed
    if ( m_hWindow )
    {
        KillTimer( m_hWindow, m_uTimerID );  // Call is OK if no timer
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

//
// Assignment operator - not allowed so no implementation
//

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Get the size of space allocated for the image
//
//  Parameters:
//      pImageSize - receives the size
//
//  Returns:
//      void
//==========================================================================//
void ImageButton::GetImageSize( SIZE* pImageSize ) const
{
    if ( pImageSize == nullptr )
    {
        throw ParameterException( L"pImageSize", __FUNCTION__ );
    }
    memcpy( pImageSize, &m_ImageSize, sizeof ( m_ImageSize ) );
}

//===============================================================================================//
//  Description:
//      Evaluate the preferred size of this window
//
//  Parameters:
//      pPreferredSize - receives the preferred size
//
//  Returns:
//      void
//===============================================================================================//
void ImageButton::GetPreferredSize( SIZE* pPreferredSize )
{
    HDC     hdc        = nullptr;
    HFONT   hfont      = nullptr;
    HGDIOBJ oldFont    = nullptr;
    LPCWSTR pszCaption = nullptr;
    SIZE    captionSize= { 0, 0 };
    LOGFONT logFont;

    if ( pPreferredSize == nullptr )
    {
        throw ParameterException( L"pPreferredSize ", __FUNCTION__ );
    }
    pPreferredSize->cx = m_ImageSize.cx;
    pPreferredSize->cy = m_ImageSize.cy;

    // Get the line height, use the desktop
    pszCaption = m_StaticText.GetText().c_str();
    if ( m_bDisplayCaption && pszCaption )
    {
        memset( &logFont, 0, sizeof ( logFont ) );
        m_StaticText.GetFont().GetLogFont( &logFont );
        hdc = GetDC( nullptr );
        if ( hdc )
        {
            /// No throwing until clean up
            hfont = CreateFontIndirect( &logFont );
            if ( hfont )
            {
                oldFont = SelectObject( hdc, hfont );
                GetTextExtentPoint32( hdc, pszCaption, lstrlen( pszCaption ), &captionSize );
                // Reset
                if ( oldFont )
                {
                    SelectObject( hdc, oldFont );
                }
                DeleteObject( hfont );
            }
            ReleaseDC( nullptr, hdc );
        }
    }

    // Horizontal width
    pPreferredSize->cx = 2;          // Left border
    if ( m_bDisplayCaption )
    {
        pPreferredSize->cx += PXSMaxInt( m_ImageSize.cx, captionSize.cx );
    }
    else
    {
        pPreferredSize->cx += m_ImageSize.cx;
    }
    pPreferredSize->cx += 2;    // Right border

    // Vertical height
    pPreferredSize->cy += 2;                 // Top border
    pPreferredSize->cy += m_ImageSize.cy;    // Image height
    if ( m_bDisplayCaption )
    {
        pPreferredSize->cy += captionSize.cy;
    }
    pPreferredSize->cy += 2;   // Bottom border
}

//===============================================================================================//
//  Description:
//      Get the text for this button
//
//  Parameters:
//      None
//
//  Returns:
//      Constant Reference to string.
//===============================================================================================//
const String& ImageButton::GetText() const
{
    return m_StaticText.GetText();
}

//===============================================================================================//
//  Description:
//      Gets if the button should display a caption
//
//  Parameters:
//      None
//
//  Returns:
//      true if caption is to be displayed, else false
//===============================================================================================//
bool ImageButton::IsCaptionDisplayed() const
{
    return m_bDisplayCaption;
}

//===============================================================================================//
//  Description:
//      Set the bitmap for the button
//
//  Parameters:
//      resourceID   - resource ID of the default bitmap
//      transparent  - a nominated transparent colour, should be a solid colour
//
//  Returns:
//      void
//===============================================================================================//
void ImageButton::SetBitmap( WORD resourceID, COLORREF transparent )
{
    String    ErrorDetails;
    HBITMAP   hBitmap = nullptr;
    Formatter Format;
    HINSTANCE hInstance = GetModuleHandle( nullptr );

    hBitmap = LoadBitmap( hInstance, MAKEINTRESOURCE( resourceID ) );
    if ( hBitmap == nullptr )
    {
        ErrorDetails = Format.StringUInt32( L"LoadBitmap, resourceID=%%1", resourceID );
        throw SystemException( GetLastError(), ErrorDetails.c_str(), __FUNCTION__);
    }

    try
    {
        MakeBitmaps( hBitmap, transparent );
    }
    catch ( const Exception& )
    {
        DeleteObject( hBitmap );
        throw;
    }
    DeleteObject( hBitmap );

    if ( IsEnabled() )
    {
        m_StaticImage.SetBitmap( m_hBitmapDefault );
    }
    else
    {
        m_StaticImage.SetBitmap( m_hBitmapOff );
    }
    Repaint();
}

//===============================================================================================//
//  Description:
//      Sets if the button should display a caption
//
//  Parameters:
//      displayCaption - flag to set for caption to be displayed
//
//  Returns:
//      void
//===============================================================================================//
void ImageButton::SetDisplayCaption( bool displayCaption )
{
    m_bDisplayCaption = displayCaption;
}

//===============================================================================================//
//  Description:
//      En/disabled this button
//
//  Parameters:
//      enabled - flag to set window enabled or disabled
//
//  Returns:
//      Reference to this window
//===============================================================================================//
void ImageButton::SetEnabled( bool enabled )
{
    if ( enabled )
    {
        EnableWindow( m_hWindow, TRUE );
        m_StaticImage.SetBitmap( m_hBitmapDefault );
        m_StaticImage.SetIcon( m_hIconDefault );
    }
    else
    {
        EnableWindow( m_hWindow, FALSE );
        m_StaticImage.SetBitmap( m_hBitmapOff );
        m_StaticImage.SetIcon( m_hIconOff );
    }
    m_StaticImage.SetEnabled( enabled );
    m_StaticText.SetEnabled( enabled );

    Repaint();
}

//===============================================================================================//
//  Description:
//      Set the bitmap for this image button as a filled shape of the
//      specified fill colour and border.
//
//  Parameters:
//      fill   - the colour to fill the interior of the border
//      border - the border colour, 1 pixel wide
//      shape  - defined shape constant
//
//  Remarks:
//      Must have created the window
//      Default, highlight and disabled bitmaps are created
//
//  Returns:
//      void
//===============================================================================================//
void ImageButton::SetFilledBitmap( COLORREF fill, COLORREF border, DWORD shape )
{
    HDC     hdc     = nullptr;
    HBITMAP hBitmap = nullptr;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    hdc = GetDC( m_hWindow );
    if ( hdc == nullptr )
    {
        throw SystemException( GetLastError(), L"hdc", __FUNCTION__ );
    }

    // Catch exceptions to release the DC
    try
    {
        hBitmap = PXSCreateFilledBitmap( hdc,
                                         m_ImageSize.cx, m_ImageSize.cy, fill, border, shape );
        MakeBitmaps( hBitmap, CLR_INVALID );
    }
    catch ( const Exception& )
    {
        DeleteObject( hBitmap );
        ReleaseDC( m_hWindow, hdc );
        throw;
    }
    DeleteObject( hBitmap );
    ReleaseDC( m_hWindow, hdc );

    if ( IsEnabled() )
    {
        m_StaticImage.SetBitmap( m_hBitmapDefault );
    }
    else
    {
        m_StaticImage.SetBitmap( m_hBitmapOff );
    }
    Repaint();
}


//===============================================================================================//
//  Description:
//      Set the icons for the button
//
//  Parameters:
//      defaultID - resource ID of the default icon
//      onID      - resource ID of the highlighted icon
//      offID     - resource ID of the disabled icon
//
//  Returns:
//      void
//===============================================================================================//
void ImageButton::SetIcons( WORD defaultID, WORD onID, WORD offID )
{
    HINSTANCE hInstance = GetModuleHandle( nullptr );

    // Any existing icons do not need to be freed
    m_hIconDefault = LoadIcon( hInstance, MAKEINTRESOURCE( defaultID ) );
    m_hIconOn      = LoadIcon( hInstance, MAKEINTRESOURCE( onID ) );
    m_hIconOff     = LoadIcon( hInstance, MAKEINTRESOURCE( offID ) );

    // Set the state
    if ( IsEnabled() )
    {
        m_StaticImage.SetIcon( m_hIconDefault );
    }
    else
    {
        m_StaticImage.SetIcon( m_hIconOff );
    }
    Repaint();
}

//===============================================================================================//
//  Description:
//      Set the text for this button
//
//  Parameters:
//      Text - pointer to text
//
//  Returns:
//      void
//===============================================================================================//
void ImageButton::SetText( const String& Text )
{
    m_StaticText.SetSingleLineText( Text );
    Repaint();
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Handle WM_LBUTTONDOWN event.
//
//  Parameters:
//      point - a POINT specifying the cursor position in the client area
//      keys  - which virtual keys are down
//
//  Returns:
//      void
//===============================================================================================//
void ImageButton::MouseLButtonDownEvent( const POINT& point, WPARAM /* keys */ )
{
    POINT    screen = point;
    ToolTip* pTip   = nullptr;

    m_bMouseWasPressed = true;

    // Forward event to tool tip
    pTip = ToolTip::GetInstance();
    if ( pTip )
    {
        ClientToScreen( m_hWindow, &screen );
        pTip->MouseButtonDown( screen );
    }
    Repaint();
}

//===============================================================================================//
//  Description:
//      Handle WM_LBUTTONUP event.
//
//  Parameters:
//      point - a POINT specifying the cursor position in the client area
//
//  Returns:
//      void
//===============================================================================================//
void ImageButton::MouseLButtonUpEvent( const POINT& point )
{
    UINT   code;
    HWND   hListener;
    RECT   clientRect = { 0, 0, 0, 0 };
    WPARAM wParam;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    GetClientRect( m_hWindow, &clientRect );
    if ( PtInRect( &clientRect, point ) )
    {
        m_bMouseWasPressed = false;

        // Send a BN_CLICKED message to parent
        PostMessage( GetParent( m_hWindow ), WM_COMMAND, (WPARAM)BN_CLICKED, (LPARAM)m_hWindow );

        // See if we need to inform a listener
        hListener = GetAppMessageListener();
        if ( hListener )
        {
            code   = GetAppMessageCode();
            wParam = MAKEWPARAM( code, 0 );
            PostMessage( hListener, WM_COMMAND, wParam, (LPARAM)m_hWindow );
        }
    }
    Repaint();
}

//===============================================================================================//
//  Description:
//      Handle WM_MOUSEMOVE event.
//
//  Parameters:
//      point - a POINT specifying the cursor position in the client area
//      keys  - which virtual keys are down
//
//  Returns:
//      void
//===============================================================================================//
void ImageButton::MouseMoveEvent( const POINT& /* point */, WPARAM /* keys */ )
{
    RECT windowRect = { 0, 0, 0, 0 };

    SetCursor( LoadCursor( nullptr, IDC_ARROW ) );
    String TipText;

    // If this window is not contained in the foreground window
    // then ignore rest of processing
    if ( IsChild( GetForegroundWindow(), m_hWindow ) == 0 )
    {
        return;
    }

    if ( m_bMouseInside == false )
    {
        m_bMouseInside = true;

        m_StaticImage.SetBitmap( m_hBitmapOn );
        m_StaticImage.SetIcon( m_hIconOn );
        SetTimer( m_hWindow, m_uTimerID, 100, nullptr );
        Repaint();

        // Set the tool tip, using the entire button's window
        // as the owner rectangle for the tool tip
        ToolTip* pTip = ToolTip::GetInstance();
        if ( pTip )
        {
            if ( GetWindowRect( m_hWindow, &windowRect ) )
            {
                TipText = GetToolTipText();
                pTip->Show( TipText, windowRect );
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Handle the WM_RBUTTONDOWN event
//
//  Parameters:
//      point - POINT where mouse was right clicked in the client area
//  Returns:
//      void
//===============================================================================================//
void ImageButton::MouseRButtonDownEvent( const POINT& point )
{
    POINT    screen = point;
    ToolTip* pTip   = nullptr;

    // Forward event to tool tip
    pTip = ToolTip::GetInstance();
    if ( pTip )
    {
        ClientToScreen( m_hWindow, &screen );
        pTip->MouseButtonDown( screen );
    }
}

//===============================================================================================//
//  Description:
//      Paint event handler
//
//  Parameters:
//      hdc - Handle to device context
//
//  Returns:
//      void
//===============================================================================================//
void ImageButton::PaintEvent( HDC hdc )
{
    Font  TextFont;
    bool  inFrame   = false;
    HWND  hWndFrame = nullptr;
    DWORD borderStyle = m_uBorderStyle;
    RECT  bounds = { 0, 0, 0, 0 }, frameBounds = { 0, 0, 0, 0 };
    POINT point  = { 0, 0, };
    ICONMETRICS im;
    StaticControl StaticBorder;

    if ( ( m_hWindow == nullptr ) || ( hdc == nullptr ) )
    {
        return;
    }

    if ( GetClientRect( m_hWindow, &bounds ) == 0 )
    {
        return;
    }
    DrawBackground( hdc );

    // Determine the border style
    if ( IsEnabled() && g_pApplication )
    {
        // Only need to draw if receiving mouse input.
        GetCursorPos( &point );
        hWndFrame = g_pApplication->GetHwndMainFrame();
        if ( hWndFrame == GetForegroundWindow() )
        {
            GetClientRect( hWndFrame, &frameBounds );
            ScreenToClient( hWndFrame, &point );
            if ( PtInRect( &frameBounds, point ) )
            {
                inFrame = true;
            }
        }
    }

    if ( inFrame )
    {
        GetCursorPos( &point );
        ScreenToClient( m_hWindow, &point );
        if ( PtInRect( &bounds, point ) )
        {
            if ( m_bMouseWasPressed )
            {
                borderStyle = PXS_SHAPE_SUNK;
            }
            else
            {
                borderStyle = PXS_SHAPE_RECTANGLE_DOTTED;
            }
        }
    }

    // Draw the border
    if ( borderStyle != PXS_SHAPE_NONE )
    {
        StaticBorder.SetBounds( bounds );
        StaticBorder.SetShape( borderStyle );
        StaticBorder.SetShapeColour( m_crForeground );
        StaticBorder.Draw( hdc );
    }

    // Draw the image
    m_StaticImage.SetPadding( 2 );
    m_StaticImage.SetBounds( bounds );
    m_StaticImage.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    m_StaticImage.SetAlignmentY( PXS_TOP_ALIGNMENT );
    m_StaticImage.Draw( hdc );

    // Draw the text
    if ( m_StaticText.GetText().c_str() )
    {
        memset( &im, 0, sizeof ( im ) );
        im.cbSize = sizeof ( im );
        if ( SystemParametersInfo( SPI_GETICONMETRICS, 0, &im, 0 ) )
        {
            TextFont.SetLogFont( &im.lfFont );
            TextFont.Create();
            m_StaticText.SetFont( TextFont );
        }
        m_StaticText.SetPadding( 2 );
        m_StaticText.SetBounds( bounds );
        m_StaticText.SetAlignmentX( PXS_CENTER_ALIGNMENT );
        m_StaticText.SetAlignmentY( PXS_BOTTOM_ALIGNMENT );
        m_StaticText.SetForeground( m_crForeground );
        m_StaticText.Draw( hdc );
    }
}

//===============================================================================================//
//  Description:
//      Handle WM_TIMER event.
//
//  Parameters:
//      timerID - The identifier of the timer that fired this event
//
//  Remarks:
//      Timer events are triggered when the mouse enters the window
//
//  Returns:
//      void.
//===============================================================================================//
void ImageButton::TimerEvent( UINT_PTR timerID )
{
    POINT cursorPos  = { 0, 0 };
    RECT  windowRect = { 0, 0, 0, 0 };

    // Determine if the cursor is still in the window
    GetCursorPos( &cursorPos );
    GetWindowRect ( m_hWindow, &windowRect );
    if ( PtInRect( &windowRect, cursorPos ) == 0 )
    {
        // Mouse is no longer inside the window, kill the timer
        // and reset the button's appearance.
        KillTimer( m_hWindow, timerID );
        m_bMouseInside     = false;
        m_bMouseWasPressed = false;
        if ( IsEnabled() )
        {
            m_StaticImage.SetBitmap( m_hBitmapDefault );
            m_StaticImage.SetIcon( m_hIconDefault );
        }
        else
        {
            m_StaticImage.SetBitmap( m_hBitmapOff );
            m_StaticImage.SetIcon( m_hIconOff );
        }
        Repaint();
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Make the default, highlighted and disabled bitmaps for this
//      image button
//
//  Parameters:
//      hBitmapDefault - the default bitmap
//      transparent    - a nominated transparent colour
//
//  Returns:
//      void
//===============================================================================================//
void ImageButton::MakeBitmaps( HBITMAP hBitmapDefault, COLORREF transparent )
{
    String    ErrorDetails;
    Formatter Format;

    // First delete any existing bitmap handles
    if ( m_hBitmapDefault )
    {
        DeleteObject( m_hBitmapDefault );
        m_hBitmapDefault = nullptr;
    }

    if ( m_hBitmapOn )
    {
        DeleteObject( m_hBitmapOn );
        m_hBitmapOn = nullptr;
    }

    if ( m_hBitmapOff )
    {
        DeleteObject( m_hBitmapOff );
        m_hBitmapOff = nullptr;
    }

    // Set new bitmaps
    if ( hBitmapDefault )
    {
        // Default bitmap
        m_hBitmapDefault = (HBITMAP)CopyImage( hBitmapDefault, IMAGE_BITMAP, 0, 0, 0 );

        // Highlighted/On
        m_hBitmapOn = (HBITMAP)CopyImage( hBitmapDefault, IMAGE_BITMAP, 0, 0, 0 );
        if ( m_hBitmapOn == nullptr )
        {
            throw SystemException( GetLastError(), L"CopyImage", __FUNCTION__ );
        }
        PXSBitmapRgbToGrb( m_hBitmapOn );

        // Disabled/Off bitmap
        m_hBitmapOff = (HBITMAP)CopyImage( hBitmapDefault, IMAGE_BITMAP, 0, 0, 0 );
        if ( m_hBitmapOn == nullptr )
        {
            throw SystemException( GetLastError(), L"CopyImage", __FUNCTION__ );
        }
        PXSBitmapToGreyScale( m_hBitmapOff );

        // If specified, make the bitmaps transparent
        if ( ( transparent    != CLR_INVALID ) &&
             ( m_crBackground != CLR_INVALID )  )
        {
            PXSReplaceBitmapColour( m_hBitmapDefault, transparent, m_crBackground );
            PXSReplaceBitmapColour( m_hBitmapOn     , transparent, m_crBackground );
            PXSReplaceBitmapColour( m_hBitmapOff    , transparent, m_crBackground );
        }
    }
}
