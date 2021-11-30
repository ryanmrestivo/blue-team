///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Window Class Implementation
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
#include "PxsBase/Header Files/Window.h"

// 2. C System Files
#include <Richedit.h>
#include <WindowsX.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateChars.h"
#include "PxsBase/Header Files/AllocateWChars.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/HandCursor.h"
#include "PxsBase/Header Files/ImageButton.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/MenuItem.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/Shell.h"
#include "PxsBase/Header Files/SystemException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////


// Default constructor
Window::Window()
       :m_bGradientVertical( false ),
        m_bUseParentRectGradient( false ),
        m_bControlsCreated( false ),
        m_uAppMessageCode( PXS_APP_MSG_NONE ),
        m_uBorderStyle( PXS_SHAPE_NONE ),
        m_hWindow( nullptr ),
        m_hWndAppMessageListener( nullptr ),
        m_hBackgroundBrush( nullptr ),
        m_crBackground( CLR_INVALID ),
        m_crBorder( CLR_INVALID ),
        m_crBorderHiLight( CLR_INVALID ),
        m_crForeground( CLR_INVALID ),
        m_crGradient1( CLR_INVALID ),
        m_crGradient2( CLR_INVALID ),
        m_crMenu( CLR_INVALID ),
        m_crMenuHiLight( CLR_INVALID ),
        HIDE_WINDOW_TIMER_ID( 1 ),
        m_recHideBitmap(),
        m_hWndHideListener( nullptr ),
        m_hBitmapHide( nullptr ),
        m_hBitmapHideOn( nullptr ),
        m_WndClassEx(),
        m_CreateStruct(),
        m_Font(),
        m_RichEditLibrary(),
        m_Statics(),
        m_bDoubleBuffered( false ),
        m_nHorizontalGap( 0 ),
        m_nVerticalGap( 0 ),
        MENU_ITEM_IMAGE_AREA_WIDTH( 22 ),
        MAX_MENU_ITEM_CHARS( 100 ),
        m_uLayoutStyle( PXS_LAYOUT_STYLE_NONE ),
        m_ToolTipText()
{
    // Class registration
    memset( &m_WndClassEx, 0, sizeof ( m_WndClassEx ) );
    m_WndClassEx.cbSize        = sizeof ( m_WndClassEx );
    m_WndClassEx.style         = CS_CLASSDC | CS_DBLCLKS;
    m_WndClassEx.lpfnWndProc   = Window::WindowProc;
    m_WndClassEx.cbClsExtra    = 0;
    m_WndClassEx.cbWndExtra    = 0;
    m_WndClassEx.hInstance     = GetModuleHandle( nullptr );
    m_WndClassEx.hIcon         = nullptr;
    m_WndClassEx.hCursor       = nullptr;
    m_WndClassEx.hbrBackground = nullptr;
    m_WndClassEx.lpszMenuName  = nullptr;
    m_WndClassEx.lpszClassName = L"PXSWindowClass";
    m_WndClassEx.hIconSm       = nullptr;

    // Creation parameters
    memset( &m_CreateStruct, 0, sizeof ( m_CreateStruct ) );
    m_CreateStruct.lpCreateParams = this;
    m_CreateStruct.hInstance      = GetModuleHandle( nullptr );
    m_CreateStruct.hMenu          = nullptr;
    m_CreateStruct.hwndParent     = nullptr;
    m_CreateStruct.cy             = 100;
    m_CreateStruct.cx             = 100;
    m_CreateStruct.y              = 0;
    m_CreateStruct.x              = 0;
    m_CreateStruct.style          = WS_CHILD | WS_VISIBLE;
    m_CreateStruct.lpszName       = nullptr;
    m_CreateStruct.lpszClass      = m_WndClassEx.lpszClassName;
    m_CreateStruct.dwExStyle      = 0;

    // Default colours
    m_crBackground    = GetSysColor( COLOR_WINDOW );
    m_crBorder        = GetSysColor( COLOR_WINDOW );
    m_crBorderHiLight = GetSysColor( COLOR_ACTIVEBORDER );
    m_crForeground    = GetSysColor( COLOR_WINDOWTEXT );
    m_crGradient1     = CLR_INVALID;
    m_crGradient2     = CLR_INVALID;
    m_crMenu          = GetSysColor( COLOR_MENU );
    m_crMenuHiLight   = GetSysColor( COLOR_MENUHILIGHT );

    // Initialise
    memset( &m_recHideBitmap, 0, sizeof ( m_recHideBitmap ) );
}

// Copy constructor - not allowed so no implementation

// Destructor
Window::~Window()
{
    // Clean up
    KillTimer( m_hWindow, HIDE_WINDOW_TIMER_ID );

    // Delete the 'close' images
    if ( m_hBitmapHide )
    {
        DeleteObject( m_hBitmapHide );
    }

    if ( m_hBitmapHideOn )
    {
        DeleteObject( m_hBitmapHideOn );
    }

    if ( m_hBackgroundBrush )
    {
        DeleteObject( m_hBackgroundBrush );
    }

    try
    {
        DestroyWindowHandle();
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
//      Append text to the window
//
//  Parameters:
//      Text - the text to append
//
//  Remarks:
//      Window must have been created before calling this method.
//      Works for Rich Edit as well.
//
//  Returns:
//      int
//===============================================================================================//
void Window::AppendText( const String& Text )
{
    int  length;
    HWND hWndPrevious;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }
    length       = GetWindowTextLength( m_hWindow );
    hWndPrevious = SetFocus( m_hWindow );
    SendMessage( m_hWindow, EM_SETSEL, (WPARAM)length, length );
    SendMessage( m_hWindow, EM_REPLACESEL, FALSE, (LPARAM)Text.c_str() );
    SendMessage( m_hWindow, EM_SETSEL, (WPARAM)-1, -1 );
    if ( hWndPrevious )
    {
        SetFocus( hWndPrevious );
    }
}

//===============================================================================================//
//  Description:
//      Create a window
//
//  Parameters:
//      pParent - pointer to the the parent window
//
//  Returns:
//      int
//===============================================================================================//
int Window::Create( const Window* pParentWindow )
{
    if ( pParentWindow == nullptr )
    {
        throw ParameterException( L"pParentWindow", __FUNCTION__ );
    }

    return Create( pParentWindow->GetHwnd() );
}

//===============================================================================================//
//  Description:
//      Create a window, register class if not already done so.
//
//  Parameters:
//      hWndParent - handle to the parent window
//
//  Returns:
//      Zero
//===============================================================================================//
int Window::Create( HWND hWndParent )
{
    DWORD      rtlStyle = 0;
    String     ClassName;
    WNDCLASSEX wcx;

    // Only allow creation once
    if ( m_hWindow )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    // If parent has Right to Left reading, propagate it to the child
    ClassName = m_WndClassEx.lpszClassName;
    if ( ClassName.CompareI( L"STATIC" ) == 0 )
    {
        if ( WS_EX_RTLREADING & GetWindowLongPtr( hWndParent, GWL_EXSTYLE ) )
        {
            m_CreateStruct.style |= SS_RIGHT;
        }
    }
    else if ( ClassName.CompareI( MSFTEDIT_CLASS ) == 0 )
    {
        // For the rich edit control, need to load the library
        // otherwise CreateWindowEx will fail
        m_RichEditLibrary.LoadSystemLibrary( L"Msftedit.dll" );
    }

    // Propagate RTL
    if ( WS_EX_RIGHT & GetWindowLongPtr( hWndParent, GWL_EXSTYLE ) )
    {
        rtlStyle |= WS_EX_RIGHT;
    }

    if ( WS_EX_RTLREADING & GetWindowLongPtr( hWndParent, GWL_EXSTYLE ) )
    {
        rtlStyle |= WS_EX_RTLREADING;
    }

    // Register the class if have not already done so.
    memset( &wcx, 0, sizeof ( wcx ) );
    wcx.cbSize = sizeof ( wcx );
    if ( GetClassInfoEx( m_WndClassEx.hInstance,
                         m_WndClassEx.lpszClassName, &wcx ) == 0 )
    {
        if ( RegisterClassEx( &m_WndClassEx ) == 0 )
        {
            throw SystemException( GetLastError(),
                                   L"RegisterClassEx", __FUNCTION__ );
        }
    }
    m_hWindow = CreateWindowEx( m_CreateStruct.dwExStyle | rtlStyle,
                                m_CreateStruct.lpszClass,
                                m_CreateStruct.lpszName,
                                static_cast<DWORD>( m_CreateStruct.style ),
                                m_CreateStruct.x,
                                m_CreateStruct.y,
                                m_CreateStruct.cx,
                                m_CreateStruct.cy,
                                hWndParent,
                                m_CreateStruct.hMenu,
                                m_CreateStruct.hInstance,
                                m_CreateStruct.lpCreateParams);

    if ( m_hWindow == nullptr )
    {
        throw SystemException( GetLastError(), L"CreateWindowEx", __FUNCTION__);
    }
    SendMessage( m_hWindow, WM_SETFONT, (WPARAM)GetStockObject( DEFAULT_GUI_FONT ), 0 );

    return 0;
}

//===============================================================================================//
//  Description:
//    Destroy the window's handle
//
//  Parameters:
//      None
//
//  Remarks:
//      DestroyWindowHandle avoids name clash with DestroyWindow.
//
//  Returns:
//      void
//===============================================================================================//
void Window::DestroyWindowHandle()
{
    if ( m_hWindow  )
    {
        // Observed an access violation in DestroyWindow if the window
        // has focus so hide it first
        ShowWindow( m_hWindow, SW_HIDE );
        DestroyWindow( m_hWindow );
        m_hWindow = nullptr;
    }
}

//===============================================================================================//
//  Description:
//      Layout this window's children
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Window::DoLayout()
{
    LayoutChildren();
    PositionHideBitmap();
    Repaint();
}

//===============================================================================================//
//  Description:
//      Yield the processor so that paint events can be processed.
//
//  Parameters:
//      timerIDEvent - the timer's identifier
//
//  Remarks:
//      Will only process WM_PAINT messages
//
//  Returns:
//      void
//===============================================================================================//
void Window::DoPaintEvents( UINT_PTR timerIDEvent )
{
    MSG msg;

    // Check class scope variables
    if ( m_hWindow == nullptr ) return;

    // Do the paint messages
    while ( IsWindowVisible( m_hWindow ) &&
            PeekMessage( &msg, nullptr, WM_PAINT, WM_PAINT, PM_REMOVE ) )
    {
        TranslateMessage( &msg );
        DispatchMessage( &msg );
    }

    // See if there are timer messages to process
    if ( timerIDEvent )
    {
        while ( IsWindowVisible( m_hWindow ) &&
                PeekMessage( &msg, nullptr, WM_TIMER, WM_TIMER, PM_REMOVE ) )
        {
            // Only process timers for this window
            if ( timerIDEvent == msg.wParam )
            {
                TranslateMessage( &msg );
                DispatchMessage( &msg );
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Set this focus on this window
//
//  Parameters:
//      None
//
//  Remarks:
//      Window must have been created
//
//  Returns:
//      void
//===============================================================================================//
void Window::Focus()
{
    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }
    SetFocus( m_hWindow );
}

//===============================================================================================//
//  Description:
//      Get the application message for this window
//
//  Parameters:
//      None
//
//  Returns:
//      the application message
//===============================================================================================//
UINT Window::GetAppMessageCode() const
{
    return m_uAppMessageCode;
}

//===============================================================================================//
//  Description:
//      Get the application message listener for this window
//
//  Parameters:
//      None
//
//  Returns:
//      HWND of this window's listener
//===============================================================================================//
HWND Window::GetAppMessageListener() const
{
    return m_hWndAppMessageListener;
}

//===============================================================================================//
//  Description:
//      Get the background colour for this window
//
//  Parameters:
//      None
//
//  Returns:
//      COLORREF of background colour.
//===============================================================================================//
COLORREF Window::GetBackground() const
{
    return m_crBackground;
}

//===============================================================================================//
//  Description:
//      Get the bounding rectangle relative to the parent
//
//  Parameters:
//      pbounds - a RECT structure to receive the bounds
//
//  Returns:
//      void
//===============================================================================================//
void Window::GetBounds( RECT *pBounds ) const
{
    if ( pBounds == nullptr )
    {
        throw ParameterException( L"pBounds", __FUNCTION__ );
    }
    memset( pBounds, 0, sizeof ( RECT ) );

    // If already created the window can get it otherwise use CREATESTRUCT
    if ( m_hWindow )
    {
        HWND  hParent = nullptr;
        POINT origin  = { 0, 0 };

        GetWindowRect( m_hWindow, pBounds );

        // Translate the screen coordinates relative to the parent
        hParent = GetParent( m_hWindow );
        ClientToScreen( hParent, &origin );
        pBounds->left   -= origin.x;
        pBounds->top    -= origin.y;
        pBounds->right  -= origin.x;
        pBounds->bottom -= origin.y;
    }
    else
    {
        pBounds->left   = m_CreateStruct.x;
        pBounds->top    = m_CreateStruct.y;
        pBounds->right  = m_CreateStruct.x + m_CreateStruct.cx;
        pBounds->bottom = m_CreateStruct.y + m_CreateStruct.cy;
    }
}

//===============================================================================================//
//  Description:
//      Get the border style for this window
//
//  Parameters:
//      None
//
//  Returns:
//      A defined constant
//===============================================================================================//
DWORD Window::GetBorderStyle() const
{
    return m_uBorderStyle;
}

//===============================================================================================//
//  Description:
//      Get the client drawable rectangle of the window.
//
//  Parameters:
//      pClientSize - a SIZE structure to receive the size
//
//  Returns:
//      void
//===============================================================================================//
void Window::GetClientSize( SIZE* pClientSize ) const
{
    if ( pClientSize == nullptr )
    {
        throw ParameterException( L"pClientSize", __FUNCTION__ );
    }

    // If already created the window can get it otherwise use CREATESTRUCT
    if ( m_hWindow )
    {
        RECT bounds = { 0, 0, 0, 0 };

        GetClientRect( m_hWindow, &bounds );
        pClientSize->cx = bounds.right  - bounds.left;
        pClientSize->cy = bounds.bottom - bounds.top;
    }
    else
    {
        pClientSize->cx = m_CreateStruct.cx;
        pClientSize->cy = m_CreateStruct.cy;
    }
}

//===============================================================================================//
//  Description:
//      Get if the window is double buffered for drawing
//
//  Parameters:
//      None
//
//  Returns:
//      true if double buffered, else false
//===============================================================================================//
bool Window::GetDoubleBuffered() const
{
    return m_bDoubleBuffered;
}

//===============================================================================================//
//  Description:
//      Get the extended style for this window
//
//  Parameters:
//      None
//
//  Returns:
//      DWORD extended style of this window
//===============================================================================================//
DWORD Window::GetExStyle() const
{
    DWORD exStyle = 0;

    // If window has been created use the API otherwise CREATESTRUCT
    if ( m_hWindow )
    {
        exStyle = (DWORD)GetWindowLongPtr( m_hWindow, GWL_EXSTYLE );
    }
    else
    {
        exStyle = m_CreateStruct.dwExStyle;
    }

    return exStyle;
}

//===============================================================================================//
//  Description:
//      Get this window's font
//
//  Parameters:
//      none
//
//  Returns:
//      constant reference to this window's font
//===============================================================================================//
const Font& Window::GetFont() const
{
    return m_Font;
}

//===============================================================================================//
//  Description:
//      Get the foreground colour for this window
//
//  Parameters:
//      None
//
//  Returns:
//      COLORREF of foreground colour.
//===============================================================================================//
COLORREF Window::GetForeground() const
{
    return m_crForeground;
}

//===============================================================================================//
//  Description:
//      Get the window handle of this window
//
//  Parameters:
//      None
//
//  Returns:
//      HWND
//===============================================================================================//
HWND Window::GetHwnd() const
{
    return m_hWindow;
}

//===============================================================================================//
//  Description:
//      Get the layout properties of this window
//
//  Parameters:
//      pLayoutStyle   - receives the named layout constant
//      pHorizontalGap - receives the horizontal layout gap
//      pVerticalGap   - receives vertical layout gap
//
//  Returns:
//      void
//===============================================================================================//
void Window::GetLayoutProperties( DWORD* pLayoutStyle, int* pHorizontalGap, int* pVerticalGap )
{
    if (  ( pLayoutStyle   == nullptr ) ||
          ( pHorizontalGap == nullptr ) ||
          ( pVerticalGap   == nullptr )  )
    {
        throw ParameterException( L"nullptr", __FUNCTION__ );
    }

    *pLayoutStyle   = m_uLayoutStyle;
    *pHorizontalGap = m_nHorizontalGap;
    *pVerticalGap   = m_nVerticalGap;
}

//===============================================================================================//
//  Description:
//      Get the position of this window relative to the parent
//
//  Parameters:
//      pLocation - a POINT structure to receive the position
//
//  Returns:
//      void
//===============================================================================================//
void Window::GetLocation( POINT* pLocation ) const
{
    RECT bounds = { 0, 0, 0, 0 };

    if ( pLocation == nullptr )
    {
        throw ParameterException( L"pLocation", __FUNCTION__ );
    }
    GetBounds( &bounds );
    pLocation->x = bounds.left;
    pLocation->y = bounds.top;
}

//===============================================================================================//
//  Description:
//      Get the size of this window.
//
//  Parameters:
//      pSize - receives the size
//
//  Returns:
//      void
//===============================================================================================//
void Window::GetSize( SIZE* pSize ) const
{
    RECT bounds = { 0, 0, 0, 0 };

    if ( pSize == nullptr )
    {
        throw ParameterException( L"pSize", __FUNCTION__ );
    }
    GetBounds( &bounds );
    pSize->cx = bounds.right  - bounds.left;
    pSize->cy = bounds.bottom - bounds.top;
}

//===============================================================================================//
//  Description:
//      Get the style for this window
//
//  Parameters:
//      None
//
//  Returns:
//      LONG style of this window
//===============================================================================================//
LONG Window::GetStyle() const
{
    LONG style = 0;

    // If window has been created use the API otherwise CREATESTRUCT
    if ( m_hWindow )
    {
        style = GetWindowLong( m_hWindow, GWL_STYLE );
    }
    else
    {
        style = m_CreateStruct.style;
    }

    return style;
}

//===============================================================================================//
//  Description:
//      Get the this window's text
//
//  Parameters:
//      pText - string object to receive the text
//
//  Returns:
//      void
//===============================================================================================//
void Window::GetText( String* pText ) const
{
    int      maxCount = 0;
    size_t   numChars = 0;
    wchar_t* pString  = nullptr;
    AllocateWChars StringBuffer;

    if ( pText == nullptr )
    {
        throw ParameterException( L"pText", __FUNCTION__ );
    }
    *pText = PXS_STRING_EMPTY;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    maxCount = GetWindowTextLength( m_hWindow );
    if ( maxCount > 0 )
    {
        maxCount = PXSAddInt32( maxCount, 1 );      // NULL terminator
        numChars = PXSCastInt32ToSizeT( maxCount );
        pString  = StringBuffer.New( numChars );
        GetWindowText( m_hWindow, pString, maxCount );
        pString[ numChars - 1 ] = PXS_CHAR_NULL;
        *pText = pString;
    }
}

//===============================================================================================//
//  Description:
//      Get the tool tip for this window
//
//  Parameters:
//      None
//
//  Returns:
//      String object of tool tip text
//===============================================================================================//
const String& Window::GetToolTipText() const
{
    return m_ToolTipText;
}

//===============================================================================================//
//  Description:
//      Get this window's class name in the WNDCLASSEX structure
//
//  Parameters:
//      pClassName - receives the class name
//
//  Returns:
//      void
//===============================================================================================//
void Window::GetWndClassExClassName( String* pClassName ) const
{
    if ( pClassName == nullptr )
    {
        throw ParameterException( L"pClassName", __FUNCTION__ );
    }
    *pClassName = m_WndClassEx.lpszClassName;
}

//===============================================================================================//
//  Description:
//      Determine if window is enabled
//
//  Parameters:
//      None
//
//  Returns:
//      true if window enabled, else false.
//===============================================================================================//
bool Window::IsEnabled() const
{
    if ( m_hWindow && IsWindowEnabled( m_hWindow ) )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Determine if window is right to left reading
//
//  Parameters:
//      None
//
//  Returns:
//      true if window is RTL otherwise false.
//===============================================================================================//
bool Window::IsRightToLeftReading()
{
    bool rtlReading = false;

    if ( m_hWindow )
    {
        if ( WS_EX_RTLREADING & GetWindowLongPtr( m_hWindow, GWL_EXSTYLE ) )
        {
            rtlReading = true;
        }
    }
    else
    {
        // Window not yet created
        if ( WS_EX_RTLREADING & m_CreateStruct.dwExStyle )
        {
            rtlReading = true;
        }
    }

    return rtlReading;
}

//===============================================================================================//
//  Description:
//      Determine if window is visible or not
//
//  Parameters:
//      None
//
//  Remarks
//      Note, if a window's parent is not visible, then the windows itself
//      is considered not visible regardless of the value of its WS_VISIBLE
//      style bit.
//
//  Returns:
//      true if window visible, else false.
//===============================================================================================//
bool Window::IsVisible() const
{
    bool visible = false;  // Assume not visible

    // If window already exists get its property otherwise use CREATESTRUCT
    if ( m_hWindow )
    {
        if ( IsWindowVisible( m_hWindow ) )
        {
            visible = true;
        }
    }
    else
    {
        if ( WS_VISIBLE & m_CreateStruct.style )
        {
            visible = true;
        }
    }

    return visible;
}

//===============================================================================================//
//    Description:
//        Transform the specified rectangle from this window's
//      coordinates to those of the screen.
//
//    Parameters:
//        pBounds - on input contains the rectangle relavtive to this
//                  window, on output contains screen coordinates
//
//    Returns:
//        void
//===============================================================================================//
void Window::RectangleToScreen( RECT* pBounds ) const
{
    int   width, height;
    POINT screenPosition;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    if ( pBounds == nullptr )
    {
        throw ParameterException( L"pBounds", __FUNCTION__ );
    }

    // Transform
    screenPosition.x = pBounds->left;
    screenPosition.y = pBounds->top;
    if ( ClientToScreen( m_hWindow, &screenPosition ) == 0 )
    {
        throw SystemException( GetLastError(), L"ClientToScreen", __FUNCTION__ );
    }
    width  = pBounds->right  - pBounds->left;
    height = pBounds->bottom - pBounds->top;
    pBounds->left  = screenPosition.x;
    pBounds->right = pBounds->left + width;
    pBounds->top   = screenPosition.y;
    pBounds->bottom= pBounds->top + height;
}

//===============================================================================================//
//  Description:
//      Redraw the entire window and its children immediately
//
//  Parameters:
//      None
//
//  Returns:
//      void.
//===============================================================================================//
void Window::RedrawAllNow() const
{
    if ( m_hWindow == nullptr )
    {
        return;
    }
    RedrawWindow( m_hWindow, nullptr, nullptr, RDW_UPDATENOW | RDW_ALLCHILDREN );
}

//===============================================================================================//
//  Description:
//      Repaint entire window, erase background
//
//  Parameters:
//      None
//
//  Returns:
//      void.
//===============================================================================================//
void Window::Repaint() const
{
    if ( m_hWindow == nullptr )
    {
        return;
    }
    InvalidateRect( m_hWindow, nullptr, FALSE );
    UpdateWindow( m_hWindow );
}

//===============================================================================================//
//  Description:
//      Redraw the entire window and its children immediately
//
//  Parameters:
//      None
//
//  Returns:
//      void.
//===============================================================================================//
void Window::RepaintChildren() const
{
    HWND hWndChild;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    hWndChild = GetWindow( m_hWindow, GW_CHILD );
    while ( hWndChild )
    {
        InvalidateRect( hWndChild, nullptr, FALSE );
        UpdateWindow( hWndChild );
        hWndChild = GetWindow( hWndChild, GW_HWNDNEXT );
    }
}

//===============================================================================================//
//  Description:
//      Mirror this window's position for right to left reading
//
//  Parameters:
//      parentWidth - the width of the window's parent in pixels
//
//  Returns:
//      void
//===============================================================================================//
void Window::RtlMirror( int parentWidth )
{
    int  width  = 0;
    RECT bounds = { 0, 0, 0, 0 };

    GetBounds( &bounds );
    width = bounds.right - bounds.left;
    bounds.right = parentWidth  - bounds.left;
    bounds.left  = bounds.right - width;
    SetBounds( bounds );
}

//===============================================================================================//
//  Description:
//      Mirror this window's static controls for right to left reading
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Window::RtlStatics()
{
    int    width  = 0;
    RECT   bounds = { 0, 0, 0, 0 };
    size_t i = 0;
    size_t numStatics = m_Statics.GetSize();
    SIZE   clientSize = { 0, 0 };
    StaticControl Static;

    GetClientSize( &clientSize );
    for ( i = 0; i < numStatics; i++ )
    {
        Static = m_Statics.Get( i );
        Static.GetBounds( &bounds );
        width  = bounds.right - bounds.left;
        bounds.right = clientSize.cx  - bounds.left;
        bounds.left  = bounds.right - width;
        Static.SetBounds( bounds );
        if ( PXS_LEFT_ALIGNMENT == Static.GetAlignmentX() )
        {
            Static.SetAlignmentX( PXS_RIGHT_ALIGNMENT );
        }
        else if ( PXS_RIGHT_ALIGNMENT == Static.GetAlignmentX() )
        {
            Static.SetAlignmentX( PXS_LEFT_ALIGNMENT );
        }
        m_Statics.Set( i, Static );
    }
}

//===============================================================================================//
//  Description:
//      Scroll to the botton of the window
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void Window::ScrollToBottom()
{
    int minPos = 0, maxPos = 0;
    LRESULT length;

    if ( m_hWindow == nullptr )
    {
        return;     // Nothing to do
    }
    GetScrollRange( m_hWindow, SB_VERT, &minPos, &maxPos );
    SetScrollPos( m_hWindow, SB_VERT, maxPos, TRUE );
    SendMessage( m_hWindow, WM_VSCROLL, MAKEWPARAM( SB_BOTTOM, 0 ), 0 );

    // Set the caret at the end
    if ( IsWindowEnabled( m_hWindow ) )
    {
        length = SendMessage( m_hWindow, WM_GETTEXTLENGTH, 0, 0 );
        SendMessage( m_hWindow, EM_SETSEL, (WPARAM)length, length );
    }
}

//===============================================================================================//
//  Description:
//      Scroll to to the top of the window
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Window::ScrollToTop()
{
    if ( m_hWindow == nullptr )
    {
        return;     // Nothing to do
    }
    SendMessage( m_hWindow, EM_SETSEL, 0, 0 );
}

//===============================================================================================//
//  Description:
//      Set the application message for this window
//
//  Parameters:
//      appMessageCode - application code for this window
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetAppMessageCode( UINT appMessageCode )
{
    if ( appMessageCode <= WM_APP )
    {
        throw ParameterException( L"appMessageCode", __FUNCTION__ );
    }
    m_uAppMessageCode = appMessageCode;
}

//===============================================================================================//
//  Description:
//      Set the application message listener for this window
//
//  Parameters:
//      hWndAppMessageListener - this window's listener
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetAppMessageListener( HWND hWndAppMessageListener )
{
    m_hWndAppMessageListener = hWndAppMessageListener;
}

//===============================================================================================//
//  Description:
//      Set the background colour for this window
//
//  Parameters:
//      background - colour Reference of the background colour
//
//  Returns:
//    void
//===============================================================================================//
void Window::SetBackground( COLORREF background )
{
    m_crBackground = background;
    if ( m_hBackgroundBrush )
    {
        DeleteObject( m_hBackgroundBrush );
        m_hBackgroundBrush = nullptr;
    }

    if ( m_crBackground != CLR_INVALID )
    {
        m_hBackgroundBrush = CreateSolidBrush( m_crBackground );
    }
}

//===============================================================================================//
//  Description:
//      Set the background gradient properties
//
//  Parameters:
//      gradient1        - RGB colour of start of gradient
//      gradient2        - RGB colour of end of gradient
//      gradientVertical - true if gradient is vertical, else false
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetBackgroundGradient( COLORREF gradient1,
                                    COLORREF gradient2,  bool gradientVertical )
{
    m_crGradient1            = gradient1;
    m_crGradient2            = gradient2;
    m_bGradientVertical      = gradientVertical;
}

//===============================================================================================//
//  Description:
//      Set the border style for this window
//
//  Parameters:
//      borderStyle - a defined shape constant
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetBorderStyle( DWORD borderStyle )
{
    m_uBorderStyle = borderStyle;
    Repaint();
}

//===============================================================================================//
//  Description:
//      Set the bounds for this window relative to the parent.
//
//  Parameters:
//      bounds - then new bounding rectangle of window in pixels
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetBounds( const RECT& bounds )
{
    return SetBounds( bounds.left,
                      bounds.top, bounds.right - bounds.left, bounds.bottom - bounds.top);
}

//===============================================================================================//
//  Description:
//      Set the bounds for this window relative to the parent.
//
//  Parameters:
//      x      - the x position
//      y      - the y position
//      width  - the width
//      height - the height
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetBounds( int x, int y, int width, int height )
{
    // If already created the window can get it otherwise use CREATESTRUCT
    if ( m_hWindow )
    {
        MoveWindow( m_hWindow, x, y, width, height, TRUE );
    }
    else
    {
        m_CreateStruct.x  = x;
        m_CreateStruct.y  = y;
        m_CreateStruct.cx = width;
        m_CreateStruct.cy = height;
    }
}

//===============================================================================================//
//  Description:
//      Set this window's colours
//
//  Parameters:
//      background  - the background colour
//      foreground  - the foreground colour
//      menuHiLight - the menu highlight colour
//
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetColours( COLORREF background, COLORREF foreground, COLORREF menuHiLight )
{
    m_crBackground  = background;
    m_crForeground  = foreground;
    m_crMenuHiLight = menuHiLight;
}

//===============================================================================================//
//  Description:
//      Set the bitmaps to hide this window
//
//  Parameters:
//      hideListener - window listening for messages from the hide button
//      hideID       - resource identifier for the bitmap
//      hideOnID     - resource identifier for the high-lighted bitmap
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetHideBitmaps( HWND hideListener, WORD hideID, WORD hideOnID )
{
    HMODULE hModule = GetModuleHandle( nullptr );

    if ( hideListener == nullptr )
    {
        throw ParameterException( L"hideListener", __FUNCTION__ );
    }
    m_hWndHideListener = hideListener;

    // Clean up
    if ( m_hBitmapHide )
    {
        DeleteObject( m_hBitmapHide );
        m_hBitmapHide = nullptr;
    }

    if ( m_hBitmapHideOn )
    {
        DeleteObject( m_hBitmapHideOn );
        m_hBitmapHideOn = nullptr;
    }

    // Load
    if ( hideID )
    {
        m_hBitmapHide = LoadBitmap( hModule, MAKEINTRESOURCE( hideID ) );
        if ( m_hBitmapHide == nullptr )
        {
            throw SystemException( GetLastError(), L"LoadBitmap", __FUNCTION__ );
        }
    }

    if ( hideOnID )
    {
        m_hBitmapHideOn = LoadBitmap( hModule, MAKEINTRESOURCE( hideOnID ) );
        if ( m_hBitmapHideOn == nullptr )
        {
            throw SystemException( GetLastError(), L"LoadBitmap", __FUNCTION__ );
        }
    }
}

//===============================================================================================//
//  Description:
//      Put the specified rich text on to the clip board
//
//  Parameters:
//      RichText - the text
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetClipboardRichText( const String& RichText )
{
    UINT      format;
    size_t    numBytes = 0;
    char*     pszAnsi  = nullptr;
    Formatter Format;
    AllocateChars AnsiString;

    if ( RichText.IsNull() )        // Allow setting ""
    {
        return;                     // Nothing to do
    }

    // Register
    format = RegisterClipboardFormat( CF_RTF );
    if ( format == 0 )
    {
        throw SystemException( GetLastError(), L"RegisterClipboardFormat", __FUNCTION__ );
    }

    // Convert to ANSI. Note, not escaping Unicode characters as this should
    // already have been done.
    numBytes = RichText.GetAnsiMultiByteLength();
    pszAnsi  = AnsiString.New( numBytes );
    Format.StringToAnsi( RichText, pszAnsi, numBytes );
    CopyDataToClipboard( format, pszAnsi, numBytes );
}

//===============================================================================================//
//  Description:
//      Put the specified text on to the clipboard
//
//  Parameters:
//      Text - the text
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetClipboardText( const String& Text )
{
    size_t numBytes = Text.GetLength();

    if ( Text.IsNull() )    // Allow setting ""
    {
        return;             // Nothing to do
    }

    // Convert to ANSI.
    numBytes = PXSMultiplySizeT( numBytes, sizeof ( wchar_t ) );
    numBytes = PXSAddSizeT( numBytes, sizeof ( wchar_t ) );  // Terminator
    CopyDataToClipboard( CF_UNICODETEXT, Text.c_str(), numBytes );
}

//===============================================================================================//
//  Description:
//      Set if the window is double buffered for drawing
//
//  Parameters:
//      doubleBuffered - flag to indicate if drawing used buffered graphics
//
//  Returns:
//      Reference to this window
//===============================================================================================//
void Window::SetDoubleBuffered( bool doubleBuffered )
{
    m_bDoubleBuffered = doubleBuffered;
}

//===============================================================================================//
//  Description:
//      Set the window to en/disabled
//
//  Parameters:
//      enabled - flag to set window enabled or disabled
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetEnabled( bool enabled )
{
    if ( m_hWindow == nullptr )
    {
        return;
    }

    if ( enabled )
    {
        EnableWindow( m_hWindow, TRUE );
    }
    else
    {
        EnableWindow( m_hWindow, FALSE );
    }
    Repaint();
}

//===============================================================================================//
//  Description:
//      Add/remove an extended style to/from a window
//
//  Parameters:
//      exStyle - an extended window style setting
//      add     - if true added to existing style, else it is removed
//
//  Remarks:
//      20140306:
//      On fail of SetWindowLong, log rather than throw. Oleg R. reported
//      application aborted when starting. This method is called to detect
//      RTL reading at programme start up.
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetExStyle( DWORD exStyle, bool add )
{
    LONG  result     = 0;
    DWORD lastError  = 0;
    DWORD newExStyle = GetExStyle();
    String    Message;
    Formatter Format;

    if ( add )
    {
        newExStyle |= exStyle;
    }
    else
    {
        newExStyle &= ~exStyle;
    }

    // If window already exists set it otherwise use CREATESTRUCT
    if ( m_hWindow )
    {
        SetLastError( ERROR_SUCCESS );
        result    = SetWindowLong( m_hWindow, GWL_EXSTYLE, static_cast<LONG>( newExStyle ) );
        lastError = GetLastError();
        if ( ( result == 0 ) && ( lastError != ERROR_SUCCESS ) )
        {
            // Log and continue
            Message  = L"SetWindowLong failed: newExStyle = ";
            Message += Format.UInt32( newExStyle );
            Message += L", SetWindowLong";
            PXSLogSysError( lastError, Message.c_str() );
        }

        // Refresh window's data, ignore errors
        SetWindowPos( m_hWindow,
                      nullptr,
                      0,
                      0,
                      0,
                      0,
                      SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
    }
    else
    {
        m_CreateStruct.dwExStyle = newExStyle;
    }
}

//===============================================================================================//
//  Description:
//      Set this window's font
//
//  Parameters:
//      FontObject - the font
//
//  Remarks:
//      Window must have been created
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetFont( const Font& FontObject )
{
    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    m_Font = FontObject;
    SendMessage( m_hWindow, WM_SETFONT, (WPARAM)m_Font.GetHandle(), MAKELPARAM( TRUE, 0 ) );
}

//===============================================================================================//
//  Description:
//      Set the foreground colour for this window
//
//  Parameters:
//      foreground - colour Reference of the foreground colour
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetForeground( COLORREF foreground )
{
    m_crForeground = foreground;
    Repaint();
}

//===============================================================================================//
//  Description:
//      Set the layout properties of this window
//
//  Parameters:
//      layoutStyle   - one of a set of named layout constants
//      horizontalGap - the horizontal layout gap
//      verticalGap   - the vertical layout gap
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetLayoutProperties( DWORD layoutStyle, int horizontalGap, int verticalGap )
{
    m_uLayoutStyle   = layoutStyle;
    m_nHorizontalGap = horizontalGap;
    m_nVerticalGap   = verticalGap;
    if ( m_hWindow )
    {
        DoLayout();
    }
}

//===============================================================================================//
//  Description:
//      Set the position of this window relative to the parent
//
//  Parameters:
//      x - x position of the window
//      y - y position of the window
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetLocation( int x, int y )
{
    SIZE size = { 0, 0 };

    GetSize( &size);
    SetBounds( x, y, size.cx, size.cy );
}
//===============================================================================================//
//  Description:
//      Set the location of this window relative to the parent
//
//  Parameters:
//      location - a POINT structure holding new position of the window
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetLocation( const POINT& location )
{
    return SetLocation( location.x, location.y );
}

//===============================================================================================//
//  Description:
//      Set Right to Left Reading
//
//  Parameters:
//      rightToLeftReading - true if want RTL otherwise false
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetRightToLeftReading( bool rightToLeftReading )
{
    HWND    hWndChild;
    LONG    exStyle   = 0, style = 0;
    LONG    rtlStyle  = WS_EX_RIGHT | WS_EX_RTLREADING;
    LONG    rtlButton = BS_LEFTTEXT | BS_RIGHT;
    wchar_t szClassName[ MAX_PATH + 1 ] = { 0 };

    SetExStyle( WS_EX_RIGHT | WS_EX_RTLREADING, rightToLeftReading );
    if ( WS_VSCROLL & GetStyle() )
    {
        SetExStyle( WS_EX_LEFTSCROLLBAR ,  rightToLeftReading );
        SetExStyle( WS_EX_RIGHTSCROLLBAR, !rightToLeftReading );
    }

    // Set child window style bits
    if ( m_hWindow == nullptr )
    {
        return;     // No children
    }

    hWndChild = GetWindow( m_hWindow, GW_CHILD );
    while ( hWndChild )
    {
        // Set extended style
        exStyle = GetWindowLong( hWndChild, GWL_EXSTYLE );
        if ( rightToLeftReading )
        {
            exStyle |= rtlStyle;
        }
        else
        {
            exStyle &= ~rtlStyle;
        }
        SetWindowLong( hWndChild, GWL_EXSTYLE, exStyle );

        // Set style specific bits depending on the window class
        memset( szClassName, 0, sizeof ( szClassName ) );
        if ( GetClassName( hWndChild, szClassName, ARRAYSIZE( szClassName ) ) )
        {
            szClassName[ ARRAYSIZE( szClassName ) - 1 ] = PXS_CHAR_NULL;
            style = GetWindowLong( hWndChild, GWL_STYLE );
            if ( PXSCompareString( szClassName, L"STATIC", false ) == 0 )
            {
                // Note, SS_LEFT = 0 so need to set its bit
                if ( rightToLeftReading )
                {
                    style |= SS_RIGHT;
                }
                else
                {
                    style &= ~SS_RIGHT;
                }
            }
            else if ( PXSCompareString( szClassName, L"BUTTON", false ) == 0)
            {
                if ( rightToLeftReading )
                {
                    style |= rtlButton;
                }
                else
                {
                    style &= ~rtlButton;
                }
            }
            SetWindowLong( hWndChild, GWL_STYLE, style );
        }

        // Refresh the window's data
        SetWindowPos( hWndChild,
                      nullptr,
                      0,
                      0,
                      0,
                      0,
                      SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED );
        hWndChild = GetWindow( hWndChild, GW_HWNDNEXT );
    }
}

//===============================================================================================//
//  Description:
//      Set the size of this window.
//
//  Parameters:
//      width  - new width
//      height  - new height
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetSize( int width, int height )
{
    RECT bounds = { 0, 0, 0, 0 };

    GetBounds( &bounds);
    SetBounds( bounds.left, bounds.top, width, height );
}

//===============================================================================================//
//  Description:
//      Set the size of this window.
//
//  Parameters:
//      size - a SIZE structure holding new dimension of the window
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetSize( const SIZE& size )
{
    SetSize( size.cx, size.cy );
}

//===============================================================================================//
//  Description:
//      Add/remove a style to/from a window
//
//  Parameters:
//      style - a window style setting
//      add   - if true add to the style, else it is removed
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetStyle( LONG style, bool add )
{
    LONG  result    = 0, newStyle = 0;
    DWORD lastError = 0;

    newStyle = GetStyle();
    if ( add )
    {
        newStyle |= style;
    }
    else
    {
        newStyle &= ~style;
    }

    // If window already exists set it, otherwise use CREATESTRUCT
    if ( m_hWindow )
    {
        SetLastError( ERROR_SUCCESS );
        result    = SetWindowLong( m_hWindow, GWL_STYLE, newStyle );
        lastError = GetLastError();
        if ( ( result == 0 ) && ( lastError != ERROR_SUCCESS ) )
        {
            throw SystemException( lastError, L"GWL_STYLE", "SetWindowLong" );
        }

        // Refresh window's data, ignore errors
        SetWindowPos( m_hWindow,
                      nullptr,
                      0,
                      0,
                      0,
                      0,
                      SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_FRAMECHANGED);
    }
    else
    {
        m_CreateStruct.style = newStyle;
    }
}

//===============================================================================================//
//  Description:
//      Set the text for this window.
//
//  Parameters:
//      Text - the text
//
//  Remarks:
//      Window needs to be created first otherwise has no effect
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetText( const String& Text )
{
    BOOL result = 0;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    if ( Text.c_str() )
    {
        result = SetWindowText( m_hWindow, Text.c_str() );
    }
    else
    {
        result = SetWindowText( m_hWindow, PXS_STRING_EMPTY );
    }
    if ( result == 0 )
    {
        throw SystemException( GetLastError(), L"SetWindowText", __FUNCTION__ );
    }
    Repaint();
}

//===============================================================================================//
//  Description:
//      Set the tool tip for this window
//
//  Parameters:
//      ToolTipText - tool tip text string
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetToolTipText( const String& ToolTipText )
{
    m_ToolTipText = ToolTipText;
}

//===============================================================================================//
//  Description:
//      Sets this window to the top in the z-order
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetToTopZ()
{
    if ( m_hWindow == nullptr )
    {
        return;     // Nothing to do
    }
    SetWindowPos( m_hWindow,
                  nullptr,  // = HWND_TOP,
                  0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW );
}

//===============================================================================================//
//  Description:
//      Sets a window to in/visible.
//
//  Parameters:
//        visible - flag to set window visible or hidden
//
//  Returns:
//      void
//===============================================================================================//
void Window::SetVisible( bool visible )
{
    // If window already exists get its property otherwise use CREATESTRUCT
    if ( m_hWindow )
    {
        if ( visible )
        {
            ShowWindow( m_hWindow, SW_SHOW );
        }
        else
        {
            ShowWindow( m_hWindow, SW_HIDE );
        }
    }
    else
    {
        if ( visible )
        {
            m_CreateStruct.style |= WS_VISIBLE;
        }
        else
        {
            m_CreateStruct.style &= ~WS_VISIBLE;
        }
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Handle WM_ACTIVATE event.
//
//  Parameters:
//      activated - true is this window is being activated, else false
//
//  Returns:
//      void.
//===============================================================================================//
void Window::ActivateEvent( bool /* activated */ )
{
}

//===============================================================================================//
//  Description:
//        Handle application specific events, i.e. above WM_APP.
//
//  Parameters:
//        uMsg   - the message number, should be greater than WM_APP
//        wParam - message specific data
//        lParam - the HWND of the window that sent the message
//
//  Returns:
//      void
//===============================================================================================//
void Window::AppMessageEvent( UINT   /* uMsg */,
                              WPARAM /* wParam */, LPARAM /* lParam */ )
{
}

//===============================================================================================//
//  Description:
//      Handle WM_CTLCOLORSTATIC event.
//
//  Parameters:
//      None
//
//  Returns:
//      LRESULT of handle to a brush that the system uses to paint the control
//===============================================================================================//
LRESULT Window::ColorStaticEvent( HWND hWndStatic )
{
    LRESULT result = 0;

    if ( hWndStatic == nullptr )
    {
        return 0;
    }

    if ( WS_EX_TRANSPARENT & GetWindowLongPtr( hWndStatic, GWL_EXSTYLE ) )
    {
        result = reinterpret_cast<LRESULT>( GetStockObject( HOLLOW_BRUSH ) );
    }
    else
    {
        if ( IsWindowEnabled( hWndStatic ) )
        {
            // Return the brush used for a this window's background. Note
            // this is not the background of the static but both would
            // normally be expected to have the same background.
            if ( m_hBackgroundBrush )
            {
                result = reinterpret_cast<LRESULT>( m_hBackgroundBrush );
            }
        }
        else
        {
            result = reinterpret_cast<LRESULT>(GetSysColorBrush(COLOR_BTNFACE));
        }
    }

    return result;
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
void Window::CloseEvent()
{
    DestroyWindowHandle();
}

//===============================================================================================//
//  Description:
//      Handle WM_COMMAND events.
//
//  Parameters:
//      wParam - The low-order word of wParam identifies the command ID
//               of the menu item, control, or accelerator. The high-order
//               word of wParam specifies the notification message if the
//               message is from a control. If the message is from an
//               accelerator, the high-order word is 1. If the message is from
//                a menu, the high-order word is 0.
//
//      lParam - Identifies the control that sends the message if the message
//               is from a control. Otherwise, lParam is 0.
//
//  Returns:
//      0 if handled, else non-zero.
//===============================================================================================//
LRESULT Window::CommandEvent( WPARAM /* wParam */, LPARAM /* lParam */ )
{
    return 0;
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
void Window::CopyDataEvent( HWND /* hWnd */, const COPYDATASTRUCT* /* pCds */ )
{
}

//===============================================================================================//
//  Description:
//      Handle WM_DESTROY event
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Window::DestroyEvent()
{
}

//===============================================================================================//
//  Description:
//      Handle WM_DISPLAYCHANGE event for this window.
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Window::DisplayChangeEvent()
{
}

//===============================================================================================//
//  Description:
//      Draw the background for this window.
//
//  Parameters:
//      hdc - handle to the device context
//
//  Returns:
//      Non-zero if background is erased, otherwise zero.
//===============================================================================================//
void Window::DrawBackground( HDC hdc )
{
    RECT   bounds = { 0, 0, 0, 0 };
    StaticControl Static;

    if ( hdc == nullptr )
    {
        return;
    }

    // Gradient
    GetClientRect( m_hWindow, &bounds );
    Static.SetBackground( m_crBackground );
    if ( ( m_crGradient1 != CLR_INVALID ) && ( m_crGradient2 != CLR_INVALID ) )
    {
        // Reverse the horizontal gradient for Right to Left Reading
        if ( ( m_bGradientVertical    == false ) &&
             ( IsRightToLeftReading() == true  )  )
        {
            Static.SetBackgroundGradient( m_crGradient2, m_crGradient1, m_bGradientVertical );
        }
        else
        {
            Static.SetBackgroundGradient( m_crGradient1, m_crGradient2, m_bGradientVertical );
        }

        // Use the size of the parent window to draw the gradient so that
        // the child's and parent's shadings look the same
        if ( m_bUseParentRectGradient )
        {
            HWND hWndParent = GetParent( m_hWindow );
            if ( hWndParent )
            {
                GetClientRect( hWndParent, &bounds );
            }
        }
    }
    Static.SetBounds( bounds );
    Static.Draw( hdc );
}

//===============================================================================================//
//  Description:
//      Draw the close bitmap on the tab
//
//  Parameters:
//      hDC     - handle to the device context
//      tabRect - the tab's bounds
//
//  Returns:
//     void
//===============================================================================================//
void Window::DrawHideBitmap( HDC hDC )
{
    bool  obscured    = false;
    HWND  hWndChild   = nullptr;
    RECT  clientRect  = { 0, 0, 0, 0 };
    POINT cursor      = { 0, 0 };
    POINT topLeft     = { m_recHideBitmap.left , m_recHideBitmap.top    };
    POINT topRight    = { m_recHideBitmap.right, m_recHideBitmap.top    };
    POINT bottomLeft  = { m_recHideBitmap.left , m_recHideBitmap.bottom };
    POINT bottomRight = { m_recHideBitmap.right, m_recHideBitmap.bottom };
    StaticControl Static;

    if ( ( hDC == nullptr )       ||
         ( m_hWindow == nullptr ) ||
         ( m_hBitmapHide == nullptr ) )
    {
        return;     // Nothing to do
    }

    // Determine if the entire bitmap can be displayed
    ClientToScreen( m_hWindow, &topLeft     );
    ClientToScreen( m_hWindow, &topRight    );
    ClientToScreen( m_hWindow, &bottomLeft  );
    ClientToScreen( m_hWindow, &bottomRight );
    hWndChild = GetWindow( m_hWindow, GW_CHILD );
    while ( hWndChild && ( obscured == false ) )
    {
        GetWindowRect( hWndChild, &clientRect );
        if ( IsWindowVisible( hWndChild ) )
        {
            if( PtInRect( &clientRect, topLeft     ) ||
                PtInRect( &clientRect, topRight    ) ||
                PtInRect( &clientRect, bottomLeft  ) ||
                PtInRect( &clientRect, bottomRight )  )
            {
                obscured = true;
            }
        }
        hWndChild = GetWindow( hWndChild, GW_HWNDNEXT );
    }

    if ( obscured )
    {
        return;
    }

    // Highlighted image of selected tab on mouse over
    if ( ( GetCursorPos( &cursor )  ) &&
         ( ScreenToClient( m_hWindow, &cursor ) ) &&
         ( PtInRect( &m_recHideBitmap, cursor ) ) )
    {
        Static.SetBitmap( m_hBitmapHideOn );
    }
    else
    {
        Static.SetBitmap( m_hBitmapHide );
    }
    Static.SetBounds( m_recHideBitmap );
    Static.SetPadding( 0 );
    Static.Draw( hDC );
}

//===============================================================================================//
//  Description:
//      Handle WM_DRAWITEM event for menu items
//
//  Parameters:
//      pDraw - pointer to menu item data structure
//
//  Returns:
//      void
//===============================================================================================//
void Window::DrawMenuItemEvent(  const DRAWITEMSTRUCT* pDraw )
{
    bool      menuBarItem = false;
    MenuItem* pMenuItem   = nullptr;
    StaticControl Static;

    if ( pDraw == nullptr  )
    {
        return;
    }

    // NULL means the item is a separator
    pMenuItem = reinterpret_cast<MenuItem*>( pDraw->itemData );
    if ( pMenuItem )
    {
        // Background
        menuBarItem = pMenuItem->IsMenuBarItem();
        DrawMenuItemBackGround( pDraw, menuBarItem );

        // Contents
        if ( menuBarItem == false )
        {
            DrawPopupMenuImageArea( pDraw, pMenuItem->GetBitmap() );
        }
        DrawMenuItemText( pDraw, pMenuItem );
    }
    else
    {
        DrawPopupMenuSeparator( pDraw );
    }
}

//===============================================================================================//
//  Description:
//      Handle WM_ERASEBKGND event for this window.
//
//  Parameters:
//      hdc - handle to the device context
//
//  Returns:
//      Non-zero if background is erased, otherwise zero.
//===============================================================================================//
LRESULT Window::EraseBackgroundEvent( HDC /* hdc */ )
{
    return 1;       // = Erased
}

//===============================================================================================//
//  Description:
//      Handle WM_SETFOCUS event for this window
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void Window::GainFocusEvent()
{
    Repaint();
}

//===============================================================================================//
//  Description:
//      Handle WM_GETDLGCODE event for this window
//
//  Parameters:
//      None
//
//  Remarks:
//      Trap this event for controls on dialog boxes to override
//      the default dialog box behaviour when certain keys are pressed
//      e.g. arrow and tab keys.
//
//  Returns:
//      UINT of dialog code, default action is to return zero.
//===============================================================================================//
UINT Window::GetDlgCodeEvent()
{
    return 0;
}

//===============================================================================================//
//  Description:
//      Handle the WM_HELP message
//
//  Parameters:
//      pHelpInfo - pointer to a HELPINFO structure
//
//  Returns:
//      void
//===============================================================================================//
void Window::HelpEvent( const HELPINFO* pHelpInfo )
{
    // Must have allocated the application object
    if ( g_pApplication == nullptr )
    {
        return;
    }
    HWND hWndMainFrame = g_pApplication->GetHwndMainFrame();

    // Forward to the applications' main frame
    if ( hWndMainFrame )
    {
        SendMessage( hWndMainFrame, WM_HELP, 0, (LPARAM)pHelpInfo );
    }
}

//===============================================================================================//
//  Description:
//      Handle WM_KEYDOWN  event for this window.
//
//  Parameters:
//      virtKey - signed integer of virtual key code
//
//  Remarks:
//      This event is raised when a non-system key is pressed.
//      A non-system key is a key that is pressed when the ALT
//      key is not pressed or a keyboard key that is pressed
//      when a window has the keyboard focus
//
//  Returns:
//      LRESULT, 0 if it the message was processed else 1.
//===============================================================================================//
LRESULT Window::KeyDownEvent( WPARAM /* virtKey */ )
{
    return 1;   // Not processed so return 1;
}

//===============================================================================================//
//  Description:
//      Layout this window's children
//
//  Parameters:
//      None
//
//  Remarks:
//      Does an immediate repaint to ensure all children are drawn properly
//
//  Returns:
//      void
//===============================================================================================//
void Window::LayoutChildren()
{
    bool rtlReading;
    int  W = 0, H   = 0, xPos = 0;
    RECT windowRect = { 0, 0, 0, 0 };
    RECT clientRect = { 0, 0, 0, 0 };
    HWND hWndChild  = nullptr;

    if ( m_hWindow == nullptr )
    {
        return;     // Nothing to do
    }
    rtlReading = IsRightToLeftReading();
    GetClientRect( m_hWindow, &clientRect );

    // Start layout according to horizontal and vertical gaps
    if ( rtlReading )
    {
        xPos = clientRect.right - m_nHorizontalGap;
    }
    else
    {
        xPos = m_nHorizontalGap;
    }

    switch ( m_uLayoutStyle )
    {
        // Fall through
        default:
        case PXS_LAYOUT_STYLE_NONE:
            break;

        case PXS_LAYOUT_STYLE_ROW_LEFT:

            // Every thing on one row, left aligned
            hWndChild = GetWindow( m_hWindow, GW_CHILD );
            while ( hWndChild )
            {
                GetWindowRect( hWndChild, &windowRect );
                W = windowRect.right  - windowRect.left;
                H = windowRect.bottom - windowRect.top;
                MoveWindow( hWndChild, xPos, m_nVerticalGap, W, H, FALSE );
                xPos += ( W + m_nHorizontalGap );
                hWndChild = GetWindow( hWndChild, GW_HWNDNEXT );
            }
            break;

        case PXS_LAYOUT_STYLE_ROW_MIDDLE:

            // Every thing on one row, middle vertical aligned, left aligned
            // otherwise right aligned for right to left reading
            hWndChild = GetWindow( m_hWindow, GW_CHILD );
            while ( hWndChild )
            {
                GetWindowRect( hWndChild, &windowRect );
                W = windowRect.right  - windowRect.left;
                H = windowRect.bottom - windowRect.top;
                if ( rtlReading )
                {
                    xPos -= W;
                }
                if ( rtlReading )
                {
                    // RTL layout does not paint child windows correctly
                    // so need to invalidate which negates the WS_CLIPCHILDREN
                    // style, i.e.may experience flicker.
                    SetWindowPos( hWndChild,
                                  nullptr,
                                  xPos,
                                  ( clientRect.bottom - H ) / 2, 0, 0, SWP_NOZORDER | SWP_NOSIZE );
                    InvalidateRect( hWndChild, nullptr, TRUE );
                    xPos -= m_nHorizontalGap;
                }
                else
                {
                    SetWindowPos( hWndChild,
                                  nullptr,
                                  xPos,
                                  ( clientRect.bottom - H ) / 2, 0, 0,
                                  SWP_NOZORDER | SWP_NOSIZE | SWP_NOREDRAW );
                    xPos += ( W + m_nHorizontalGap );
                }
                hWndChild = GetWindow( hWndChild, GW_HWNDNEXT );
            }
            break;
    }
}

//===============================================================================================//
//  Description:
//      Handle WM_KILLFOCUS event for this window
//
//  Parameters:
//      hWnd - handle to window receiving keyboard/mouse focus.
//
//  Returns:
//      void
//===============================================================================================//
void Window::LostFocusEvent( HWND /* hWnd */ )
{
    Repaint();
}

//===============================================================================================//
//  Description:
//      Handle WM_MEASUREITEM event for menu items
//
//  Parameters:
//      pMeasure - pointer to menu item data structure
//
//  Returns:
//      void
//===============================================================================================//
void Window::MeasureMenuItemEvent( MEASUREITEMSTRUCT* pMeasure )
{
    int       charLength = 0;
    HDC       hdc        = nullptr;
    HFONT     hFont      = nullptr;
    HGDIOBJ   oldFont    = nullptr;
    MenuItem* pMenuItem  = nullptr;
    LPCWSTR   pszLabel   = nullptr;
    SIZE      textExtent = { 0, 0 };
    String    Label;
    NONCLIENTMETRICS ncm;

    if ( pMeasure == nullptr )
    {
        return;
    }

    // Get a pointer to the item's data, its a ULONG_PTR
    pMenuItem = reinterpret_cast<MenuItem*>( pMeasure->itemData );
    if ( pMenuItem == nullptr )
    {
        // Is a separator, use a nominal width and height 2
        pMeasure->itemWidth   = 50;
        pMeasure->itemHeight  =  2;
        return;
    }

    // Truncate the caption
    Label = pMenuItem->GetLabel();
    Label.Truncate( MAX_MENU_ITEM_CHARS );
    pszLabel = Label.c_str();
    if ( pszLabel == nullptr )
    {
        return;     // Nothing to measure
    }
    charLength = lstrlen( pszLabel );

    hdc = GetDC( m_hWindow );
    if ( hdc == nullptr )
    {
        return;     // Nothing to measure on
    }

    // Get the font for the menu
    memset( &ncm, 0, sizeof ( ncm ) );
    ncm.cbSize = sizeof ( ncm );
    SystemParametersInfo( SPI_GETNONCLIENTMETRICS, 0, &ncm, 0 );
    hFont = CreateFontIndirect( &ncm.lfMenuFont );
    if ( hFont )
    {
        oldFont = SelectObject( hdc, hFont );
    }
    if ( GetTextExtentPoint32( hdc, pszLabel, charLength, &textExtent ) )
    {
        pMeasure->itemWidth  = PXSCastLongToUInt32( textExtent.cx );
        pMeasure->itemWidth += 10;  // Padding
        pMeasure->itemHeight = PXSCastLongToUInt32( textExtent.cy );
    }

    if ( oldFont ) SelectObject( hdc, oldFont );
    if ( hFont   ) DeleteObject( hFont );
    ReleaseDC( m_hWindow, hdc );

    // Add a text padding of 2
    pMeasure->itemWidth  += 4;
    pMeasure->itemHeight += 4;

    // Add space for a bitmap 16*16 if its on a pop-up menu
    if ( !pMenuItem->IsMenuBarItem() )
    {
        pMeasure->itemWidth += 16;
        pMeasure->itemWidth += 4;   // Padding
        pMeasure->itemWidth += 1;   // Add space for a vertical line
        pMeasure->itemHeight = PXSMaxUInt32( pMeasure->itemHeight, 20 );
    }
}

//===============================================================================================//
//    Description:
//        Handle WM_MENUCHAR event for menu items
//
//    Parameters:
//      chUser - character entered by user
//        hMenu  - handle to menu
//
//  Remarks:
//      Menu mnemonics are not available with owner-drawn menus so need to
//      handle this message
//
//    Returns:
//        LRESULT according to WM_MENUCHAR
//===============================================================================================//
LRESULT Window::MenuCharEvent( wchar_t chUser, HMENU hMenu )
{
    LRESULT   result = MNC_IGNORE;
    int       i = 0 , count = 0;
    UINT      uItem = 0;
    wchar_t   szMnemonic[ 3 ] = { 0 };    // i.e. ampersand + char + null
    String    Label;
    MenuItem* pMenuItem = nullptr;
    MENUITEMINFO mii;

    if ( ( chUser == PXS_CHAR_NULL ) || ( hMenu == nullptr ) )
    {
        return MNC_IGNORE;
    }

    // Set the mnemonic to search for
    szMnemonic[ 0 ] = '&';
    szMnemonic[ 1 ] = chUser;
    szMnemonic[ 2 ] = PXS_CHAR_NULL;

    count = GetMenuItemCount( hMenu );
    if ( count == -1 )
    {
        return MNC_IGNORE;  // Error so ignore this event
    }

    // Cycle through the menu items and find the matching index
    while ( ( i < count ) && ( result == MNC_IGNORE ) )
    {
        memset( &mii , 0 , sizeof ( mii ) );
        mii.cbSize = sizeof ( mii );
        mii.fMask  = MIIM_DATA;
        mii.fType  = MFT_OWNERDRAW;
        uItem = PXSCastInt32ToUInt32( i );
        if ( GetMenuItemInfo( hMenu, uItem, TRUE, &mii ) )
        {
            // Note, dwItemData is of type LPTSTR
            pMenuItem = reinterpret_cast<MenuItem*>( mii.dwItemData );
            if ( pMenuItem )
            {
                // Look for the mnemonic
                Label = pMenuItem->GetLabel();
                if ( Label.IndexOfI( szMnemonic ) != PXS_MINUS_ONE )
                {
                    result = MAKELRESULT( uItem, MNC_EXECUTE );
                }
            }
        }
        i++;
    }

    return result;
}

//===============================================================================================//
//  Description:
//      Handle WM_LBUTTONDBLCLK event.
//
//  Parameters:
//      point - a POINT specifying the cursor position in the client area
//
//  Returns:
//      void
//===============================================================================================//
void Window::MouseLButtonDblClickEvent( const POINT& /* point */ )
{
}

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
void Window::MouseLButtonDownEvent(const POINT& /* point */, WPARAM /* keys */)
{
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
void Window::MouseLButtonUpEvent( const POINT& point )
{
    RECT   bounds = { 0, 0, 0, 0 };
    size_t i = 0, numStatics = m_Statics.GetSize();
    StaticControl Static;

    if ( m_hWindow == nullptr ) return;

    // See if clicked on the close button
    if ( PtInRect( &m_recHideBitmap, point ) )
    {
        SendMessage( m_hWndHideListener, PXS_APP_MSG_HIDE_WINDOW, 0, (LPARAM)m_hWindow );
    }

    // Scan for hyper-links
    for ( i = 0; i < numStatics; i++ )
    {
        Static = m_Statics.Get( i );
        if ( Static.IsHyperlink() )
        {
            Static.GetBounds( &bounds );
            if ( PtInRect( &bounds, point ) )
            {
                Shell Browser;
                Browser.NavigateToUrl( Static.GetText() );
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Handle WM_MOUSEMOVE event.
//
//  Parameters:
//      point - a POINT specifying the cursor position in the client area
//      keys  - which virtual keys are down
//
//  Remarks:
//      This method is called when the mouse is moving in a window that is:
//       - enabled;
//       - a foreground window or the child of foreground window
//
//  Returns:
//      void
//===============================================================================================//
void Window::MouseMoveEvent( const POINT& point, WPARAM /* keys */ )
{
    RECT   bounds = { 0, 0, 0, 0 };
    size_t i = 0, numStatics = m_Statics.GetSize();
    StaticControl Static;

    if ( m_hWindow == nullptr ) return;

    SetCursor( LoadCursor( nullptr, IDC_ARROW ) );

    // Handle redraw the close button if mouse is inside the bitmap
    if ( m_hBitmapHide )
    {
        if ( PtInRect( &m_recHideBitmap, point ) )
        {
            SetTimer( m_hWindow, HIDE_WINDOW_TIMER_ID, 100, nullptr );
        }
        InvalidateRect( m_hWindow, &m_recHideBitmap, FALSE );
    }

    // Scan for hyper-links to set a hand cursor
    for ( i = 0; i < numStatics; i++ )
    {
        Static = m_Statics.Get( i );
        if ( Static.IsHyperlink() )
        {
            Static.GetBounds( &bounds );
            if ( PtInRect( &bounds, point ) )
            {
                HandCursor Hand;
                Hand.Set();
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Handle the WM_RBUTTONDOWN event
//
//  Parameters:
//        point - POINT where mouse was right clicked in the client area
//  Returns:
//      void
//===============================================================================================//
void Window::MouseRButtonDownEvent( const POINT& /* point */ )
{
}

//===============================================================================================//
//  Description:
//      Handle WM_RBUTTONUP event.
//
//  Parameters:
//      point - a POINT specifying the cursor position in the client area
//
//  Returns:
//      void
//===============================================================================================//
void Window::MouseRButtonUpEvent( const POINT& /* point */ )
{
}

//===============================================================================================//
//  Description:
//      Handle WM_MOUSEWHEEL event.
//
//  Parameters:
//      keys  - virtual key flags
//      delta - delta movement of wheel
//      point - a POINT specifying the cursor position in the client area
//
//  Returns:
//      void
//===============================================================================================//
void Window::MouseWheelEvent( WORD /* keys */,
                              SHORT /* delta */, const POINT& /* point */ )
{
}

//===============================================================================================//
//  Description:
//      Handle WM_NCPAINT event.
//
//  Parameters:
//      hdc - Handle to device context
//
//  Returns:
//      void
//===============================================================================================//
void Window::NonClientPaintEvent( HDC /* hdc */ )
{
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
void Window::NotifyEvent( const NMHDR* /* pNmhdr */ )
{
}

//===============================================================================================//
//  Description:
//      Handle WM_PAINT event.
//
//  Parameters:
//      hdc - Handle to device context
//
//  Returns:
//      void
//===============================================================================================//
void Window::PaintEvent( HDC hdc )
{
    size_t i = 0, numStatics = m_Statics.GetSize();
    RECT   bounds = { 0, 0, 0, 0 };
    StaticControl Static;

    if ( ( m_hWindow == nullptr ) || ( hdc == nullptr ) )
    {
        return;     // Nothing to draw;
    }
    DrawBackground( hdc );

    // Draw static Controls
    for ( i = 0; i < numStatics; i++)
    {
        Static = m_Statics.Get( i );
        Static.Draw( hdc );
    }
    DrawHideBitmap( hdc );

    // Draw border - use a system colour because will show if user changes
    // system settings
    if ( m_uBorderStyle != PXS_SHAPE_NONE )
    {
        Static.Reset();
        GetClientRect( m_hWindow, &bounds );
        Static.SetBounds( bounds );
        Static.SetShape( m_uBorderStyle );
        Static.SetShapeColour( GetSysColor( COLOR_3DDKSHADOW ) );
        Static.Draw( hdc );
    }
}

//===============================================================================================//
//  Description:
//      Position this window's close button, if it has one
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Window::PositionHideBitmap()
{
    int    borderWidth = 0;
    SIZE   clientSize  = { 0, 0 };
    BITMAP bmHide;

    if ( ( m_hWindow == nullptr ) || ( m_hBitmapHide == nullptr ) )
    {
        return;
    }

    memset( &bmHide, 0, sizeof ( bmHide ) );
    if ( GetObject( m_hBitmapHide, sizeof ( bmHide ), &bmHide ) == 0 )
    {
        return;
    }

    // Account for a possible border, 1 pixel wide
    if ( m_uBorderStyle != PXS_SHAPE_NONE )
    {
        borderWidth = 1;
    }

    // Put the close button in a top corner
    GetClientSize( &clientSize );
    if ( clientSize.cx > bmHide.bmWidth )
    {
        // To the left for LTR and right for RTL
        m_recHideBitmap.left = clientSize.cx - bmHide.bmWidth - borderWidth;
        if ( IsRightToLeftReading() )
        {
            m_recHideBitmap.left = borderWidth;
        }
        m_recHideBitmap.top    = borderWidth;
        m_recHideBitmap.right  = m_recHideBitmap.left + bmHide.bmWidth;
        m_recHideBitmap.bottom = m_recHideBitmap.top  + bmHide.bmHeight;
    }
    else
    {
        // Not enough space to show the bitmap
        memset( &m_recHideBitmap, 0, sizeof ( m_recHideBitmap ) );
    }
}

//===============================================================================================//
//  Description:
//      Handle WM_HSCROLL and WM_VSCROLL  events.
//
//  Parameters:
//      vertical    - if true this is a vertical scroll event, else horizontal
//      scrollCode  - a scroll-bar code that indicates the type of movement
//
//  Returns:
//      LRESULT, 0 if it the message was processed else 1.
//===============================================================================================//
LRESULT Window::ScrollEvent( bool /* vertical */, int /* scrollCode */ )
{
    return 1;
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
void Window::SettingChangeEvent()
{
    DoLayout();
}

//===============================================================================================//
//  Description:
//      Handle WM_SIZE event.
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void Window::SizeEvent()
{
    DoLayout();
}

//===============================================================================================//
//  Description:
//      Handle WM_TIMER event.
//
//  Parameters:
//      timerID - The identifier of the timer that fired this event
//
//  Returns:
//      void
//===============================================================================================//
void Window::TimerEvent( UINT_PTR timerID )
{
    POINT cursorPos  = { 0, 0 };

    // Handle if the mouse has moved outside of the hide bitmap
    if ( ( timerID == HIDE_WINDOW_TIMER_ID ) && m_hBitmapHide )
    {
        GetCursorPos( &cursorPos );
        ScreenToClient( m_hWindow, &cursorPos );
        if ( PtInRect( &m_recHideBitmap, cursorPos ) == 0 )
        {
            KillTimer( m_hWindow, HIDE_WINDOW_TIMER_ID );
            InvalidateRect( m_hWindow, &m_recHideBitmap, FALSE );
        }
    }
}

//===============================================================================================//
//  Description:
//      Static window callback procedure.
//
//  Parameters:
//      hWnd   - Handle to the window that generated the message.
//      Msg    - Specifies the message.
//      wParam - Specifies additional message information.
//      lParam - Specifies additional message information.
//
//  Returns:
//      LRESULT, which is usually 0 if the message was processed, if not
//      pass back up to DefWindowProc
//===============================================================================================//
LRESULT Window::WindowProc( HWND hWnd, UINT uMsg, WPARAM wParam, LPARAM lParam )
{
    bool    showException = true;      // Normally want to show exceptions
    HDC     hdc     = nullptr;
    POINT   cursor  = { 0, 0 };
    Window* pWindow = reinterpret_cast<Window*>(
                                      GetWindowLongPtr( hWnd, GWLP_USERDATA ) );
    LRESULT  result  = 0;
    MEASUREITEMSTRUCT*    lpmis = nullptr;
    const DRAWITEMSTRUCT* pDraw = nullptr;

    // Only allow null window if its being created
    if ( ( pWindow == nullptr ) && ( uMsg != WM_CREATE ) )
    {
        return DefWindowProc( hWnd, uMsg, wParam, lParam );
    }

    try
    {
        switch ( uMsg )
        {
            default:

                // See if its an application specific message
                if ( uMsg >= WM_APP )
                {
                    pWindow->AppMessageEvent( uMsg, wParam, lParam );
                }
                else
                {
                    result = DefWindowProc( hWnd, uMsg, wParam, lParam );
                }
                break;

            case WM_ACTIVATE:

                if ( WA_INACTIVE == LOWORD( wParam ) )
                {
                    pWindow->ActivateEvent( false );
                }
                else
                {
                    pWindow->ActivateEvent( true );
                }
                break;

            case DM_GETDEFID:

                // Forward to the parent
                PostMessage( GetParent( hWnd ), uMsg, wParam, lParam );
                break;

            case WM_CLOSE:

                pWindow->CloseEvent();
                break;

            case WM_COMMAND:

                result = pWindow->CommandEvent( wParam, lParam );
                break;

            case WM_COPYDATA:

                pWindow->CopyDataEvent( reinterpret_cast<HWND>( wParam ),
                                        reinterpret_cast<COPYDATASTRUCT*>( lParam ) );
                break;

            case WM_CREATE:

                // Store the creation parameter as user data.
                // CREATESTRUCT::lpCreateParams is set on CreateWindowEx
                // and is equal to this
                if ( lParam )
                {
                    CREATESTRUCT* pCS = reinterpret_cast<CREATESTRUCT*>(lParam);
                    SetWindowLongPtr( hWnd,
                                      GWLP_USERDATA,
                                      reinterpret_cast<LONG_PTR>( pCS->lpCreateParams ) );
                }
                break;

            case WM_CTLCOLORSTATIC:

                // Will set the background colour permanently in the DC so
                // that the background colour of text in the static window
                // is the same as the parent. There will be flicker if the
                // parent has gradient background and resized, so any parents
                // wuth static windows should have a solid background.
                showException = false;
                if ( pWindow->m_crGradient1 != CLR_INVALID )
                {
                    SetBkMode( (HDC)wParam, TRANSPARENT );
                }
                else
                {
                    SetBkColor( (HDC)wParam, pWindow->m_crBackground );
                }
                result = pWindow->ColorStaticEvent( reinterpret_cast<HWND>( lParam ) );
                break;

            case WM_DISPLAYCHANGE:

                showException = false;
                pWindow->DisplayChangeEvent();
                break;

            case WM_DESTROY:

                pWindow->DestroyEvent();
                break;

            case WM_DRAWITEM:

                showException = false;
                pDraw = reinterpret_cast<const DRAWITEMSTRUCT*>( lParam );
                if ( pDraw && ( pDraw->CtlType == ODT_MENU ) )
                {
                    pWindow->DrawMenuItemEvent( pDraw );
                }
                result = TRUE;     // Documentation says should return TRUE
                break;

            case WM_ERASEBKGND:

                result = pWindow->EraseBackgroundEvent( (HDC)wParam );
                break;

            case WM_GETDLGCODE:

                result = static_cast<LRESULT>( pWindow->GetDlgCodeEvent() );
                break;

            case WM_HELP:

                pWindow->HelpEvent( reinterpret_cast<HELPINFO*>( lParam ) );
                break;

            case WM_HSCROLL:

                // Make sure this window has the focus
                showException = false;
                result = pWindow->ScrollEvent( false, LOWORD( wParam ) );
                break;

            case WM_KEYDOWN:

                result = pWindow->KeyDownEvent( wParam );
                if ( result )
                {
                    // Forward to the parent
                    PostMessage( GetParent( hWnd ), uMsg, wParam, lParam );
                }
                break;

            case WM_KILLFOCUS:

                // Will not show exceptions so as do not want focus to transfer
                showException = false;
                pWindow->LostFocusEvent( reinterpret_cast<HWND>( wParam ) );
                PostMessage( GetParent( hWnd ), uMsg, wParam, lParam );
                break;

            case WM_LBUTTONDBLCLK:

                cursor.x = GET_X_LPARAM( lParam );
                cursor.y = GET_Y_LPARAM( lParam );
                pWindow->MouseLButtonDblClickEvent( cursor );
                break;

            case WM_LBUTTONDOWN:

                cursor.x = GET_X_LPARAM( lParam );
                cursor.y = GET_Y_LPARAM( lParam );
                pWindow->MouseLButtonDownEvent( cursor, wParam );
                break;

            case WM_LBUTTONUP:

                cursor.x = GET_X_LPARAM( lParam );
                cursor.y = GET_Y_LPARAM( lParam );
                pWindow->MouseLButtonUpEvent( cursor );
                break;

            case WM_MEASUREITEM:

                showException = false;
                lpmis = reinterpret_cast<MEASUREITEMSTRUCT*>( lParam );
                if ( lpmis && ( lpmis->CtlType == ODT_MENU ) )
                {
                    pWindow->MeasureMenuItemEvent( lpmis );
                }
                break;

            case WM_MENUCHAR:

                result = pWindow->MenuCharEvent( LOWORD( wParam ),
                                                 reinterpret_cast<HMENU>( lParam ) );
                break;

            case WM_MOUSEMOVE:

                showException = false;
                if ( IsWindowEnabled( hWnd ) )
                {
                    cursor.x = GET_X_LPARAM( lParam );
                    cursor.y = GET_Y_LPARAM( lParam );
                    pWindow->MouseMoveEvent( cursor, wParam );
                }
                break;

            case WM_MOUSEWHEEL:

                // Mouse wheel cursor messages are in screen coordinates
                cursor.x = GET_X_LPARAM( lParam );
                cursor.y = GET_Y_LPARAM( lParam );
                ScreenToClient( hWnd, &cursor );
                pWindow->MouseWheelEvent( GET_KEYSTATE_WPARAM( wParam ),
                                          GET_WHEEL_DELTA_WPARAM( wParam ), cursor );
                break;

            case WM_NCACTIVATE:
            case WM_NCPAINT:

                // Process message normally
                showException = false;
                result = DefWindowProc( hWnd, uMsg, wParam, lParam );

                // Now draw any custom stuff in the non-client area
                hdc = GetWindowDC( hWnd );
                if ( hdc )
                {
                    pWindow->NonClientPaintEvent( hdc );
                    ReleaseDC( hWnd, hdc );
                }
                break;

            case WM_NOTIFY:

                pWindow->NotifyEvent( reinterpret_cast<NMHDR*>( lParam ) );
                break;

            case WM_RBUTTONDOWN:

                cursor.x = GET_X_LPARAM( lParam );
                cursor.y = GET_Y_LPARAM( lParam );
                pWindow->MouseRButtonDownEvent( cursor );
                break;

            case WM_RBUTTONUP:

                cursor.x = GET_X_LPARAM( lParam );
                cursor.y = GET_Y_LPARAM( lParam );
                pWindow->MouseRButtonUpEvent( cursor );
                break;

            case WM_PAINT:

                // Will not use the rcPaint member of PAINTSTRUCT as sometimes
                // this results in an incomplete paint when rcPaint does not
                // have the same dims as the window's client rectangle.
                showException = false;
                pWindow->DoPaint();
                break;

            case WM_SETFOCUS:

                showException = false;
                pWindow->GainFocusEvent();
                break;

            case WM_SETTINGCHANGE:

                pWindow->SettingChangeEvent();
                break;

            case WM_SIZE:

                pWindow->SizeEvent();
                break;

            case WM_TIMER:

                showException = false;
                pWindow->TimerEvent( wParam );
                break;

            case WM_VSCROLL:

                showException = false;
                result = pWindow->ScrollEvent( true, LOWORD( wParam ) );
                break;
        }
    }
    catch ( const Exception& e )
    {
        if ( showException && IsWindowVisible( hWnd ) )
        {
            PXSShowExceptionDialog( e, hWnd );
        }
    }

    return result;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Put the specified data on the clipboard
//
//  Parameters:
//      format   - the clipboard format, e.g. CF_TEXT
//      pData    - pointer to the data
//      numBytes - number of bytes of data
//
//  Returns:
//      void
//===============================================================================================//
void Window::CopyDataToClipboard( UINT format, const void* pData, size_t numBytes )
{
    DWORD   lastError;
    void*   pLocked;
    HGLOBAL hGlobal;

    if ( format == 0 )
    {
        throw ParameterException( L"format", __FUNCTION__ );
    }

    if ( ( pData == nullptr ) || ( numBytes == 0 ) )
    {
        return;     // Nothing to do
    }

    // Copy the data
    hGlobal = GlobalAlloc( GHND, numBytes );
    if ( hGlobal == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }
    pLocked = GlobalLock( hGlobal );
    if ( pLocked == nullptr )
    {
       lastError = GetLastError();
       GlobalFree( hGlobal );
       throw SystemException( lastError, L"GlobalLock", __FUNCTION__ );
    }
    memcpy( pLocked, pData, numBytes );   // False positive C6386 with /analyze
    GlobalUnlock( hGlobal );

    if ( OpenClipboard( m_hWindow ) == 0 )
    {
        GlobalFree( hGlobal );
        throw SystemException( GetLastError(), L"OpenClipboard", __FUNCTION__ );
    }
    EmptyClipboard();

    // When SetClipboardData succeeds the data is owned by the system so
    // must not free the handle or leave it locked
    if ( SetClipboardData( format, hGlobal ) == nullptr )
    {
        GlobalFree( hGlobal );
        CloseClipboard();
        throw SystemException( GetLastError(), L"SetClipboardData", __FUNCTION__ );
    }
    CloseClipboard();
}


//===============================================================================================//
//  Description:
//    Draw the background of a menu item
//
//  Parameters:
//      pDraw       - DRAWITEMSTRUCT
//      menuBarItem - true if the item is on the menu bar, otherwise a pop-up
//
//  Returns:
//      void
//===============================================================================================//
void Window::DrawMenuItemBackGround( const DRAWITEMSTRUCT* pDraw, bool menuBarItem )
{
    POINT cursorPos  = { 0, 0 };
    RECT  windowRect = { 0, 0, 0, 0 };
    StaticControl Static;

    if ( pDraw == nullptr )
    {
        return;     // Nothing to draw
    }

    // If the item is selected draw hi-lighted
    if (  ( pDraw->itemState & ODS_SELECTED ) &&
         !( pDraw->itemState & ODS_DISABLED )  )
    {
        Static.SetBackground( GetSysColor( COLOR_MENUHILIGHT ) );
        Static.SetShape( PXS_SHAPE_RECTANGLE_DOTTED );
        Static.SetShapeColour( m_crForeground );
    }
    else
    {
        if ( menuBarItem )
        {
            Static.SetBackground( GetSysColor( COLOR_MENUBAR ) );
            Static.SetBackgroundGradient( m_crGradient1, m_crGradient2, true );

            // If mouse is inside the item, high-light it
            if ( !( pDraw->itemState & ODS_DISABLED ) )
            {
                GetCursorPos( &cursorPos );
                GetWindowRect( m_hWindow, &windowRect );
                cursorPos.x -= windowRect.left;
                cursorPos.y -= windowRect.top;
                if ( PtInRect( &pDraw->rcItem, cursorPos ) )
                {
                    Static.SetBackground( GetSysColor( COLOR_MENUHILIGHT ) );
                    Static.SetBackgroundGradient( CLR_INVALID,
                                                  CLR_INVALID, true );
                    Static.SetShape( PXS_SHAPE_RECTANGLE_DOTTED );
                    Static.SetShapeColour( m_crForeground );
                }
            }
        }
        else
        {
            Static.SetBackground( GetSysColor( COLOR_MENU ) );

            // Reversed gradient for right-to-left
            if ( IsRightToLeftReading() )
            {
                Static.SetBackgroundGradient( m_crGradient1, m_crGradient2, false );
            }
            else
            {
                Static.SetBackgroundGradient( m_crGradient2, m_crGradient1, false );
            }
        }
    }
    Static.SetBounds( pDraw->rcItem );
    Static.Draw( pDraw->hDC );
}

//===============================================================================================//
//  Description:
//    Draw the menu item's text
//
//  Parameters:
//      pDraw - DRAWITEMSTRUCT
//
//  Returns:
//      void
//===============================================================================================//
void Window::DrawMenuItemText( const DRAWITEMSTRUCT* pDraw, const MenuItem* pMenuItem )
{
    int      count      = 0, nSavedDC;
    RECT     bounds     = { 0, 0, 0, 0 };
    UINT     textAlign  = DT_LEFT;
    HFONT    ulinedFont = nullptr;
    size_t   tabIndex   = 0, numChars = 0;
    HGDIOBJ  hFontObj   = nullptr, oldFont = nullptr;
    COLORREF textColour = CLR_INVALID;
    String   Label, Mnemonic;
    LOGFONT  logFont;

    if ( ( pDraw == nullptr ) || ( pMenuItem == nullptr )  )
    {
        return;     // Nothing to draw
    }

    // The text, truncate to maximum
    Label = pMenuItem->GetLabel();
    if ( Label.IsEmpty() )
    {
        return;     // Nothing to draw
    }

    // Colours
    nSavedDC   = SaveDC( pDraw->hDC );
    textColour = GetMenuItemTextColour( pDraw, pMenuItem );
    SetTextColor( pDraw->hDC, textColour );

    // Bounds and alignment
    bounds = pDraw->rcItem;
    if ( pMenuItem->IsMenuBarItem() )
    {
        textAlign = DT_CENTER;
    }
    else
    {
        // Menu item
        if ( IsRightToLeftReading() )
        {
            textAlign    = DT_RIGHT;
            bounds.left  += 5;
            bounds.right -= ( MENU_ITEM_IMAGE_AREA_WIDTH + 5 );
        }
        else
        {
            bounds.left   = MENU_ITEM_IMAGE_AREA_WIDTH + 5;
            bounds.right -= 5;
        }
    }

    // Set underline for hyper-links.
    if ( pMenuItem->IsHyperLink() )
    {
        hFontObj = GetCurrentObject( pDraw->hDC, OBJ_FONT );
        if ( hFontObj )
        {
            memset( &logFont, 0, sizeof ( logFont ) );
            GetObject( hFontObj, sizeof( logFont ), &logFont );
            logFont.lfUnderline = TRUE;
            ulinedFont = CreateFontIndirect( &logFont );
            if ( ulinedFont )
            {
                oldFont = SelectObject( pDraw->hDC, ulinedFont );
            }
        }
    }

    // Draw up to the first tab
    SetBkMode( pDraw->hDC, TRANSPARENT );
    tabIndex = Label.IndexOf( PXS_CHAR_TAB, 0 );  // -1 if no tab
    numChars = Label.GetLength();
    numChars = PXSMinSizeT( numChars, tabIndex );
    numChars = PXSMinSizeT( numChars, MAX_MENU_ITEM_CHARS );
    count    = PXSCastSizeTToInt32( numChars );
    DrawText( pDraw->hDC, Label.c_str(), count, &bounds, textAlign | DT_VCENTER | DT_SINGLELINE );

    // Remove underlined font for the accelerator mnemonic part
    if ( oldFont    ) SelectObject( pDraw->hDC, oldFont );
    if ( ulinedFont ) DeleteObject( ulinedFont );

    // Draw the part after the tab, if any
    if ( tabIndex != PXS_MINUS_ONE )
    {
        Label.SubString( tabIndex + 1, PXS_MINUS_ONE, &Mnemonic );

        // Swap the text alignment
        if ( textAlign == DT_LEFT )
        {
            textAlign = DT_RIGHT;
        }
        else
        {
            textAlign = DT_LEFT;
        }
        count = PXSCastSizeTToInt32( Mnemonic.GetLength() );
        DrawText( pDraw->hDC,
                  Mnemonic.c_str(), count, &bounds, textAlign | DT_VCENTER | DT_SINGLELINE );
    }
    RestoreDC( pDraw->hDC, nSavedDC );  // Reset
}

//===============================================================================================//
//  Description:
//    Draw the image area of a pop-up menu item
//
//  Parameters:
//      pDraw         - DRAWITEMSTRUCT
//      hBitmap       - the menu item's bitmap, ignored if the item is checked
//
//  Returns:
//      void
//===============================================================================================//
void Window::DrawPopupMenuImageArea( const DRAWITEMSTRUCT* pDraw, HBITMAP hBitmap )
{
    bool    rtlReading;
    int     xPos    = 0, yPos = 0;
    HPEN    hPen    = nullptr;
    RECT    bounds  = { 0, 0, 0, 0 };
    HGDIOBJ oldPen  = nullptr;

    if ( pDraw == nullptr )
    {
        return;     // Nothing to draw
    }
    rtlReading = IsRightToLeftReading();

    // Draw background
    bounds = pDraw->rcItem;
    if ( rtlReading )
    {
        bounds.left = bounds.right - MENU_ITEM_IMAGE_AREA_WIDTH;
        FillRect( pDraw->hDC, &bounds, (HBRUSH)GetStockObject( WHITE_BRUSH ) );

        // Vertical separator
        bounds.left -= 1;
        bounds.right = bounds.left + 1;
        FillRect( pDraw->hDC, &bounds, (HBRUSH)GetStockObject( BLACK_BRUSH ) );
        xPos = pDraw->rcItem.right - MENU_ITEM_IMAGE_AREA_WIDTH + 3;
    }
    else
    {
        bounds.right = bounds.left + MENU_ITEM_IMAGE_AREA_WIDTH;
        FillRect( pDraw->hDC, &bounds, (HBRUSH)GetStockObject( WHITE_BRUSH ) );

        // Vertical separator
        bounds.left += MENU_ITEM_IMAGE_AREA_WIDTH;
        bounds.right = bounds.left + 1;
        FillRect( pDraw->hDC, &bounds, (HBRUSH)GetStockObject( BLACK_BRUSH ) );
        xPos = 1;
    }

    // Apply an offset if selected
    yPos  = pDraw->rcItem.top;
    yPos += ( pDraw->rcItem.bottom - pDraw->rcItem.top - 16 ) / 2;
    if ( ( pDraw->itemState & ODS_SELECTED ) == 0 )
    {
        xPos += 1;
        yPos += 1;
    }

    // Draw the bitmap
    if ( hBitmap && !( pDraw->itemState & ODS_DISABLED ) )
    {
        PXSDrawTransparentBitmap( pDraw->hDC, hBitmap, xPos, yPos, GetSysColor( COLOR_MENU ) );
    }

    // If the item is checked, overlay a tick mark
    if ( pDraw->itemState & ODS_CHECKED )
    {
        hPen = CreatePen( PS_SOLID, 2, PXS_COLOUR_NAVY );
        if ( hPen )
        {
            oldPen = SelectObject( pDraw->hDC, hPen );

            xPos += 3;
            yPos += 8;
            MoveToEx( pDraw->hDC, xPos, yPos, nullptr );
            LineTo( pDraw->hDC, xPos + 3, yPos + 3 );
            LineTo( pDraw->hDC, xPos + 7, yPos - 4 );

            // Reset
            if ( oldPen )
            {
                SelectObject( pDraw->hDC, oldPen );
            }
            DeleteObject( hPen );
        }
    }
}

//===============================================================================================//
//  Description:
//    Draw a separator on a pop-up menu item
//
//  Parameters:
//      pDraw - DRAWITEMSTRUCT
//
//  Returns:
//      void
//===============================================================================================//
void Window::DrawPopupMenuSeparator( const DRAWITEMSTRUCT* pDraw )
{
    bool rtlReading;
    RECT bounds;
    StaticControl Static;

    if ( pDraw == nullptr )
    {
        return;     // Nothing to draw
    }
    rtlReading = IsRightToLeftReading();

    // Draw the background
    bounds = pDraw->rcItem;
    Static.SetBackground( GetSysColor( COLOR_MENU ) );
    if ( rtlReading )
    {
        Static.SetBackgroundGradient( m_crGradient1, m_crGradient2, false );
    }
    else
    {
        Static.SetBackgroundGradient( m_crGradient2, m_crGradient1, false );
    }
    Static.SetBounds( bounds );
    Static.Draw( pDraw->hDC );
    Static.Reset();

    // Draw the image area background
    bounds = pDraw->rcItem;
    if ( rtlReading )
    {
        bounds.left = bounds.right - MENU_ITEM_IMAGE_AREA_WIDTH;
    }
    else
    {
        bounds.right = MENU_ITEM_IMAGE_AREA_WIDTH;
    }
    Static.SetBounds( bounds );
    Static.SetBackground( PXS_COLOUR_WHITE );
    Static.Draw( pDraw->hDC );
    Static.Reset();

    // Draw vertical line
    bounds = pDraw->rcItem;
    if ( rtlReading )
    {
        bounds.left  = bounds.right - MENU_ITEM_IMAGE_AREA_WIDTH - 1;
        bounds.right = bounds.left;
    }
    else
    {
        bounds.left  = MENU_ITEM_IMAGE_AREA_WIDTH;
        bounds.right = bounds.left;
    }
    Static.SetBounds( bounds );
    Static.SetShape( PXS_SHAPE_LINE );
    Static.SetShapeColour( GetSysColor( COLOR_MENUTEXT ) );
    Static.Draw( pDraw->hDC );

    // Draw horizontal line
    bounds        = pDraw->rcItem;
    bounds.top    = bounds.top + ( bounds.bottom - bounds.top ) / 2;
    bounds.bottom = bounds.top;
    if ( rtlReading )
    {
        bounds.right = bounds.right - MENU_ITEM_IMAGE_AREA_WIDTH - 1;
    }
    else
    {
        bounds.left = MENU_ITEM_IMAGE_AREA_WIDTH;
    }
    Static.SetBounds( bounds );
    Static.SetShape( PXS_SHAPE_LINE );
    Static.SetShapeColour( GetSysColor( COLOR_MENUTEXT ) );
    Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    Static.Draw( pDraw->hDC );
}

//===============================================================================================//
//  Description:
//    Handle the paint operation between BeginPaint and EndPaint
//
//  Parameters:
//      ps - PAINTSTRUCT
//
//  Remarks:
//      Avoid throwing as caller needs to match BeginPaint with EndPaint
//
//  Returns:
//      void
//===============================================================================================//
void Window::DoPaint()
{
    LONG    width, height;
    RECT    recClient     = { 0, 0, 0, 0 };
    HDC     hCompatibleDC = nullptr, hDC = nullptr;
    HBITMAP hCompatibleBM = nullptr;
    HGDIOBJ oldBitmap     = nullptr;
    PAINTSTRUCT ps;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    // Will not use the rcPaint member of PAINTSTRUCT as sometimes
    // this results in an incomplete paint when rcPaint does not
    // have the same dims as the window's client rectangle.
    GetClientRect( m_hWindow, &recClient );
    width  = recClient.right  - recClient.left;
    height = recClient.bottom - recClient.top;

    memset( &ps, 0, sizeof ( ps ) );
    hDC = BeginPaint( m_hWindow, &ps );
    if ( hDC == nullptr )
    {
        PXSLogSysError( GetLastError(), L"Begin paint failed." );
        EndPaint( m_hWindow, &ps );
        return;
    }

    try
    {
        // If double buffering, create off-screen DC
        if ( GetDoubleBuffered() )
        {
            hCompatibleDC = CreateCompatibleDC( hDC );
            hCompatibleBM = CreateCompatibleBitmap( hDC, width, height );
            if ( hCompatibleDC && hCompatibleBM )
            {
                oldBitmap = SelectObject( hCompatibleDC, hCompatibleBM );
                PaintEvent( hCompatibleDC );
                BitBlt( hDC,
                        0, 0, width, height, hCompatibleDC, 0, 0, SRCCOPY );

                // Reset DC
                if ( oldBitmap )
                {
                    SelectObject( hCompatibleDC, oldBitmap );
                }
            }

            // Clean up
            if ( hCompatibleBM )
            {
                DeleteObject( hCompatibleBM );
            }
            if ( hCompatibleDC )
            {
                DeleteDC( hCompatibleDC );
            }
        }
        else
        {
            PaintEvent( hDC );  // Draw on ordinary device context
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
    EndPaint( m_hWindow, &ps );
}

//===============================================================================================//
//  Description:
//      Determine the colour text for the specified menu item
//
//  Parameters:
//      pDraw     - DRAWITEMSTRUCT structure
//      pMenuItem - the menu item
//
//  Returns:
//      void
//===============================================================================================//
COLORREF Window::GetMenuItemTextColour( const DRAWITEMSTRUCT* pDraw, const MenuItem* pMenuItem )
{
    POINT    cursorPos  = { 0, 0 };
    RECT     windowRect = { 0, 0, 0, 0 };
    BOOL     isHilited  = false;
    COLORREF textColour = PXS_COLOUR_BLACK;

    if ( ( pDraw == nullptr ) || ( pMenuItem == nullptr ) )
    {
        return false;
    }

    if ( pDraw->itemState & ODS_DISABLED )
    {
        return GetSysColor( COLOR_GRAYTEXT );
    }

    if ( pMenuItem->IsHyperLink() )
    {
        // Menu item
        if ( pDraw->itemState & ODS_SELECTED )
        {
            textColour = PXS_COLOUR_CYAN;
        }
        else
        {
            textColour = GetSysColor( COLOR_HOTLIGHT );
        }
    }
    else
    {
        // Determine if highlighted
        if ( pMenuItem->IsMenuBarItem() )
        {
            GetCursorPos( &cursorPos );
            GetWindowRect( m_hWindow, &windowRect );
            cursorPos.x -= windowRect.left;
            cursorPos.y -= windowRect.top;
            isHilited = PtInRect( &pDraw->rcItem, cursorPos );
        }
        else
        {
            if ( pDraw->itemState & ODS_SELECTED )
            {
                isHilited = TRUE;
            }
        }

        if ( isHilited )
        {
            textColour = GetSysColor( COLOR_HIGHLIGHTTEXT );
        }
        else
        {
            textColour = GetSysColor( COLOR_MENUTEXT );
        }
    }

    return textColour;
}
