///////////////////////////////////////////////////////////////////////////////////////////////////
//
//  Frame Window Class Implementation
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
#include "PxsBase/Header Files/Frame.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/StringArray.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/ToolTip.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
Frame::Frame()
      :m_AccelTable(),
       m_Title(),
       m_hIcon( nullptr ),
       m_hIconSmall( nullptr ),
       m_hWndChildLastFocus( nullptr )
{
    // CREATESTRUCT structure, initially invisible
    m_CreateStruct.style      = WS_OVERLAPPEDWINDOW | WS_CLIPCHILDREN;
    m_CreateStruct.cy         = CW_USEDEFAULT;
    m_CreateStruct.cx         = CW_USEDEFAULT;
    m_CreateStruct.y          = CW_USEDEFAULT;
    m_CreateStruct.x          = CW_USEDEFAULT;
    m_CreateStruct.hwndParent = nullptr;
}

// Copy constructor - not allowed so no implementation

// Destructor
// DestroyIcon not required for shared icons obtained by LoadIcon
Frame::~Frame()
{
    ToolTip* pTip;

    try
    {
        // A child of this frame may have created a tool tip, if so clean up
        ToolTip::IsInstanceCreated();
        {
            pTip = ToolTip::GetInstance();
            if ( pTip )
            {
                pTip->DestroyInstance();
            }
        }
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Maximise the frame
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Frame::Maximize()
{
    RECT workRect = { 0, 0, 0, 0 };

    // If window has been created
    if ( m_hWindow )
    {
        if ( IsWindowVisible( m_hWindow ) )
        {
            ShowWindow( m_hWindow, SW_MAXIMIZE );
        }
        else
        {
            // Set to the workspace area
            if ( SystemParametersInfo( SPI_GETWORKAREA, 0, &workRect, 0 ) )
            {
                if ( ( workRect.right  > workRect.left ) &&
                     ( workRect.bottom > workRect.top  )  )
                {
                    SetBounds( workRect );
                }
            }
        }
    }
    else
    {
        // Set the style bit
        SetStyle( WS_MAXIMIZE, true );
    }
}

//===============================================================================================//
//  Description:
//     Set the bounds of this frame based on a comma separate string
//
//  Parameters:
//      SetFrameBounds - string of top,left,right,bottom values in pixels
//
//  Returns:
//      void
//===============================================================================================//
void Frame::SetFrameBounds( const String& FrameBounds )
{
    RECT   bounds   = { 0, 0, 0, 0 };
    RECT   workArea = { 0, 0, 0, 0 };
    String Value;
    StringArray Values;

    SystemParametersInfo( SPI_GETWORKAREA, 0, &workArea, 0 );
    FrameBounds.ToArray( ',', &Values );
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
                    bounds.top,
                    bounds.right - bounds.left,
                    bounds.bottom - bounds.top, TRUE );
    }
    else
    {
        ShowWindow( m_hWindow, SW_MAXIMIZE );
    }
}

//===============================================================================================//
//  Description:
//      Set the icon used in the frame's title bar, if frame already
//      created then send a message
//
//  Parameters:
//        resourceID - Resource ID identifier of icon resource
//        bigIcon    - Flag to indicate size of icon, false = 16*16,
//                     true = 32*32 pixels using LoadIcon.
//
//  Remarks:
//      LoadIcon loads the icon resource only if it has not
//      been loaded; otherwise, it retrieves a handle to the
//      existing resource, so no need to clean up.
//
//  Returns:
//      void
//===============================================================================================//
void Frame::SetIconImage( WORD resourceID, bool bigIcon )
{
    // Must have created the window
    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    if ( resourceID == 0 )
    {
        return;     // Nothing to do
    }

    // No need to use DestroyIcon()
    if ( bigIcon )
    {
        m_hIcon = LoadIcon( GetModuleHandle( nullptr ), MAKEINTRESOURCE( resourceID ) );
        if ( m_hIcon == nullptr )
        {
            throw SystemException( GetLastError(), L"LoadIcon", __FUNCTION__ );
        }
        SendMessage( m_hWindow, WM_SETICON, (WPARAM)ICON_BIG, (LPARAM)m_hIcon );
    }
    else
    {
        m_hIconSmall = static_cast<HICON>( LoadImage( GetModuleHandle(nullptr),
                                           MAKEINTRESOURCE( resourceID ),
                                           IMAGE_ICON, 16, 16, 0 ) );
        if ( m_hIconSmall == nullptr )
        {
           throw SystemException( GetLastError(), L"LoadImage", __FUNCTION__ );
        }
        SendMessage( m_hWindow, WM_SETICON, (WPARAM)ICON_SMALL, (LPARAM)m_hIconSmall );
    }
}

//===============================================================================================//
//  Description:
//      Set the title used in the frame's title bar
//
//  Parameters:
//      pszTitle - Null terminated string i.e. the title
//
//  Returns:
//      void
//===============================================================================================//
void Frame::SetTitle( LPCWSTR pszTitle )
{
    // Check window handle
    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    // Ensure have non-null pointer, do not trim
    m_Title = pszTitle;
    if ( m_Title.c_str() == nullptr )
    {
        m_Title = PXS_STRING_EMPTY;
    }
    SendMessage( m_hWindow, WM_SETTEXT, 0, (LPARAM)m_Title.c_str() );
}

//===============================================================================================//
//  Description:
//      Show the frame
//
//  Parameters:
//      visible - flag to set frame visible or not
//
//  Remarks::
//      Sometimes the children of the frame are not correctly laid out
//      even though their bounding rectangles are correct. So resize the frame
//      slightly then return to original size. This would be similar to the
//      Java Window.pack() method.
//
//  Returns:
//      void.
//===============================================================================================//
void Frame::SetVisible( bool visible )
{
    if ( m_hWindow )
    {
        if ( visible )
        {
            RECT bounds = { 0, 0, 0, 0 };

            // SW_SHOWNORMAL flag will restore a minimized window
            ShowWindow( m_hWindow, SW_SHOWNORMAL );

            // Force a resize to pack the frame even if already
            // visible to over come paint problems on Win9X
            GetBounds( &bounds );
            bounds.right  += 1;
            bounds.bottom += 1;

            SetBounds( bounds );
            bounds.right  -= 1;
            bounds.bottom -= 1;
            SetBounds( bounds );

            // Bring the main window to the top.
            BringWindowToTop( m_hWindow );
            SetForegroundWindow( m_hWindow );
            if ( IsIconic( m_hWindow ) )
            {
                ShowWindow( m_hWindow, SW_RESTORE );
            }

            // Get any pop-ups
            HWND hWndPopUp = GetLastActivePopup( m_hWindow );
            if ( hWndPopUp && ( m_hWindow != hWndPopUp ) )
            {
                BringWindowToTop( hWndPopUp );
            }
        }
        else
        {
            ShowWindow( m_hWindow, SW_HIDE );
        }
    }
    else
    {
        // Window not yet created, set its style
        SetStyle( WS_VISIBLE, visible );
    }
}

//===============================================================================================//
//  Description:
//      Start the main message loop
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
DWORD Frame::StartMessageLoop()
{
    MSG    Msg;
    BOOL   bRet      = FALSE;
    DWORD  lastError = 0;
    HACCEL hAccel    = nullptr;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    bRet = GetMessage( &Msg, nullptr, 0, 0 );
    while ( bRet && ( lastError == 0 ) )
    {
        if ( bRet == -1 )
        {
            lastError = GetLastError();
        }
        else
        {
             // Send keyboard accelerators messages to this frame's WindowProc
            hAccel = m_AccelTable.GetHaccel();
            if ( m_hWindow && hAccel && m_AccelTable.IsEnabled() )
            {
                if ( TranslateAccelerator( m_hWindow, hAccel, &Msg ) == 0 )
                {
                    // Process normally
                    TranslateMessage( &Msg );
                    DispatchMessage( &Msg );
                }
            }
            else
            {
                TranslateMessage( &Msg );
                DispatchMessage( &Msg );
            }
        }
        bRet = GetMessage( &Msg, nullptr, 0, 0 );
    }

    return lastError;
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
void Frame::ActivateEvent( bool activated )
{
    HWND hWndFocus = nullptr;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    // If this window is becoming active, set the focus to the
    // child window that previously had it
    if ( activated )
    {
        if ( m_hWndChildLastFocus && IsChild(m_hWindow, m_hWndChildLastFocus) )
        {
            if ( m_hWndChildLastFocus != GetFocus() )
            {
                SetFocus( m_hWndChildLastFocus );
            }
        }
    }
    else
    {
        // Window is being de-activated, so store the child window that
        // has focus so it can be re-instated if/when this frame becomes
        // active again. Note, if the frame is being de-activated because it
        // is being minimized, GetFocus returns NULL. The HWND data of
        // the WM_ACTIVATE and KILL focus messages are also NULL.
        hWndFocus = GetFocus();
        if ( hWndFocus && IsChild( m_hWindow, hWndFocus ) )
        {
            m_hWndChildLastFocus = hWndFocus;
        }
    }
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
void Frame::DestroyEvent()
{
    PostQuitMessage( 0 );   // Quit the application
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
void Frame::NonClientPaintEvent( HDC hdc )
{
    int   itemCount = 0;
    UINT  uItem     = 0, i = 0;
    RECT  bounds    = { 0, 0, 0, 0 };
    RECT  itemRect  = { 0, 0, 0, 0 };
    RECT  windowRect= { 0, 0, 0, 0 };
    HRGN  hRectRgn  = nullptr;
    MENUBARINFO   mbi;
    StaticControl Static;

    if ( ( hdc == nullptr ) || ( m_hWindow == nullptr ) )
    {
        return;
    }

    // Draw the menu bar's background, but exclude the menu bar items
    // otherwise these will be overdrawn
    memset( &mbi, 0, sizeof ( mbi ) );
    mbi.cbSize = sizeof ( mbi );
    if ( GetMenuBarInfo( m_hWindow, OBJID_MENU, 0, &mbi ) )
    {
        // Evaluate the menu bars bounds
        bounds.left   = GetSystemMetrics( SM_CXFRAME )
                        + GetSystemMetrics( SM_CXBORDER ) - 1;
        bounds.top    = GetSystemMetrics( SM_CYFRAME )
                        + GetSystemMetrics( SM_CYCAPTION );
        bounds.right  = bounds.left + ( mbi.rcBar.right  - mbi.rcBar.left ) - 1;
        bounds.bottom = bounds.top  + ( mbi.rcBar.bottom - mbi.rcBar.top  ) + 1;

        // Exclude the the menu items
        memset( &windowRect, 0, sizeof ( windowRect ) );
        GetWindowRect( m_hWindow, &windowRect );
        hRectRgn = CreateRectRgn( bounds.left, bounds.top, bounds.right, bounds.bottom );
        SelectClipRgn( hdc, hRectRgn );
        itemCount = GetMenuItemCount( mbi.hMenu );
        if ( itemCount > 0 )
        {
            uItem = static_cast< UINT >( itemCount );
        }
        for ( i = 0; i < uItem; i++ )
        {
            memset( &itemRect, 0, sizeof ( itemRect ) );
            GetMenuItemRect( m_hWindow, mbi.hMenu, i, &itemRect );
            OffsetRect( &itemRect, -windowRect.left, -windowRect.top );
            ExcludeClipRect( hdc, itemRect.left, itemRect.top, itemRect.right, itemRect.bottom );
        }

        // Draw, catch as need to reset
        try
        {
            Static.SetBackground( GetSysColor( COLOR_MENUBAR ) );
            Static.SetBackgroundGradient( m_crGradient1, m_crGradient2, true );
            Static.SetBounds( bounds );
            Static.Draw( hdc );
        }
        catch ( const Exception& )
        { }  // Ignore

        SelectClipRgn( hdc, nullptr );
        DeleteObject( hRectRgn );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
