///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Tab Window Class Implementation
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

///////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////
// Include Files
///////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/TabWindow.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/SystemException.h"
#include "PxsBase/Header Files/ToolTip.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default Constructor
TabWindow::TabWindow()
          :HORIZ_TAB_GAP( 4 ),
           INTERNAL_LEADING( 2 ),
           BASELINE_HEIGHT( 1 ),
           MIN_TAB_WIDTH( 24 ),             // Allow room for the close button
           MIN_TAB_HEIGHT( BASELINE_HEIGHT + 24 ),
           m_nPadding( 4 ),
           m_nTabWidth( 0 ),
           m_nMaxTabWidth( 250 ),
           m_nTabStripHeight( 0 ),          // 0 implies preferred layout height
           m_uTimerID( 1001 ),
           m_uSelectedTabIndex( 0 ),
           m_uMouseInTabIndex( PXS_MINUS_ONE ),
           m_hBitmapClose( nullptr ),
           m_hBitmapCloseOn( nullptr ),
           m_crTabGradient1( CLR_INVALID ),
           m_crTabGradient2( CLR_INVALID ),
           m_crTabHiLiteBackground( CLR_INVALID ),
           m_Tabs()

{
    // Creation parameters
    m_CreateStruct.cx    = 100;
    m_CreateStruct.cy    = 20;
    m_CreateStruct.style = WS_CHILD | WS_CLIPCHILDREN | WS_VISIBLE;

    // Properties
    try
    {
        SetBackground( GetSysColor( COLOR_BTNFACE ) );
        m_crTabHiLiteBackground = GetSysColor( COLOR_3DHILIGHT );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy Constructor - not allowed so no implementation

// Destructor
TabWindow::~TabWindow()
{
    size_t i = 0, numTabs;
    TYPE_TAB_DATA TabData;

    // Delete the 'close' images
    if ( m_hBitmapClose )
    {
        DeleteObject( m_hBitmapClose );
    }

    if ( m_hBitmapCloseOn )
    {
        DeleteObject( m_hBitmapCloseOn );
    }

    // Ensure the timer has been killed
    if ( m_hWindow )
    {
        KillTimer( m_hWindow, m_uTimerID );  // Call is OK if no timer
    }

    // Free the bitmaps
    try
    {
        numTabs = m_Tabs.GetSize();
        for ( i = 0; i < numTabs; i++ )
        {
            TabData = m_Tabs.Get( i );
            if ( TabData.hBitmap )
            {
                DeleteObject( TabData.hBitmap );
            }
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

// Assignment operator - not allowed so no implementation

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
//===============================================================================================//
//  Description:
//      Add a tab, set to invisible if it is not the selected tab
//
//  Parameters:
//      hBitmap     - bitmap image to show on tab, can be NULL
//      pszName     - the text to display on the tab
//      pszToolTip  - tool tip text
//      allowClose  - indicates the tab can be closed
//      tabID       - a numberical value to identify this tab
//      pWindow     - the window to display when the tab is selected
//
//  Returns:
//      the zero based index postion of the tab
//===============================================================================================//
size_t TabWindow::Add( HBITMAP hBitmap,
                       LPCWSTR pszName,
                       LPCWSTR pszToolTip, bool allowClose, DWORD tabID, Window* pWindow )
{
    size_t    i = 0, numTabs = 0, newTabIndex = m_Tabs.GetSize();
    Formatter Format;
    TYPE_TAB_DATA TabData;

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    if ( pWindow == nullptr )
    {
        throw ParameterException( L"pWindow", __FUNCTION__ );
    }

    // Make sure not adding the same window twice
    numTabs = m_Tabs.GetSize();
    for ( i = 0; i < numTabs; i++ )
    {
        TabData = m_Tabs.Get( i );
        if ( pWindow == TabData.pWindow )
        {
            throw SystemException( ERROR_OBJECT_ALREADY_EXISTS, L"pWindow", __FUNCTION__ );
        }
    }
    memset( &TabData, 0, sizeof ( TabData ) );
    TabData.visible = TRUE;
    if ( hBitmap )
    {
        TabData.hBitmap = (HBITMAP)CopyImage( hBitmap, IMAGE_BITMAP, 0, 0, 0 );
    }

    if ( pszName )
    {
        StringCchCopy( TabData.szName, ARRAYSIZE( TabData.szName ), pszName );
    }

    if ( pszToolTip )
    {
        StringCchCopy( TabData.szToolTip, ARRAYSIZE( TabData.szToolTip ), pszToolTip );
    }

    if ( allowClose )
    {
        TabData.allowClose = TRUE;
    }
    TabData.tabID   = tabID;
    TabData.state   = PXS_STATE_UNKNOWN;
    TabData.pWindow = pWindow;
    m_Tabs.Add( TabData );

    // Select the tab after doing a layout so when it becomes visible
    // it is already at the correct location
    DoLayout();
    SetSelectedTabIndex( newTabIndex );
    Repaint();

    return newTabIndex;
}

//===============================================================================================//
//  Description:
//      Do the layout, position the tabs and the associated pages
//
//  Parameters:
//      None
//
//  Remarks:
//      Keep the layout of a tab in sync with PaintEvent
//
//  Returns:
//      Void
//===============================================================================================//
void TabWindow::DoLayout()
{
    int     tabWidth   = 0, xPos = 0, yPos = 0, tabHeight = 0;
    size_t  numTabs    = 0, i = 0;
    HDC     hDC        = nullptr;
    HFONT   hFont      = nullptr;
    HGDIOBJ hOldFont   = nullptr;
    RECT    rect = { 0, 0, 0, 0 }, clientRect = { 0, 0, 0, 0 };
    TEXTMETRIC    tm;
    TYPE_TAB_DATA TabData;

    if ( m_hWindow == nullptr )
    {
        return;     // Nothing to do
    }
    GetClientRect( m_hWindow, &clientRect );

    // Evaluate the tab height
    hDC = GetDC( m_hWindow );
    if ( hDC == nullptr )
    {
        return;
    }
    hFont = m_Font.GetHandle();
    if ( hFont )
    {
        hOldFont = SelectObject( hDC, hFont );
    }
    memset( &tm, 0, sizeof ( tm ) );
    GetTextMetrics( hDC, &tm );
    tabHeight = tm.tmHeight + tm.tmExternalLeading;
    tabHeight += ( 2 * m_nPadding );
    tabHeight = PXSMaxInt( tabHeight, MIN_TAB_HEIGHT );

    // Catch exceptions as need to release DC
    try
    {
        xPos    = 0;
        yPos    = INTERNAL_LEADING;
        numTabs = m_Tabs.GetSize();
        for ( i = 0; i < numTabs; i++ )
        {
            TabData = m_Tabs.Get( i );
            SetRect( &TabData.bounds, -1, -1, -1, -1 );   // i.e. off screen
            if ( TabData.visible )
            {
                tabWidth = EvaluateTabWidth( hDC , tabHeight, TabData );

                // Set the tab's bounds
                TabData.bounds.left   = xPos;
                TabData.bounds.right  = xPos + tabWidth;
                TabData.bounds.top    = yPos;
                TabData.bounds.bottom = yPos + tabHeight;
                if ( IsRightToLeftReading() )
                {
                    TabData.bounds.right = clientRect.right - xPos;
                    TabData.bounds.left  = TabData.bounds.right - tabWidth;
                }
                xPos += tabWidth;
            }
            m_Tabs.Set( i, TabData );
        }
        yPos += tabHeight;

        // Layout the the pages in the remaining vertical space
        if ( yPos < m_nTabStripHeight )
        {
            SetRect( &rect, 1, m_nTabStripHeight, clientRect.right, clientRect.bottom );
        }
        else
        {
            SetRect( &rect, 1, yPos, clientRect.right, clientRect.bottom );
        }
        LayoutPages( rect );
    }
    catch ( const Exception& )  // Ignore
    {}

    // Reset DC
    if ( hOldFont )
    {
        SelectObject( hDC, hOldFont );
    }
    ReleaseDC( m_hWindow, hDC );
    Repaint();
}

//===============================================================================================//
//  Description:
//      Get the bounds occupied by the tabs
//
//  Parameters:
//      pBounds - receives the bounds
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::GetTabBounds( RECT* pBounds )
{
    size_t i = 0, numTabs = m_Tabs.GetSize();
    TYPE_TAB_DATA TabData;

    if ( pBounds == nullptr )
    {
        throw ParameterException( L"pBounds", __FUNCTION__ );
    }
    memset( pBounds, 0, sizeof ( RECT ) );

    if ( numTabs == 0 )
    {
        return;
    }

    // As all tabs are on one line, the bounds are effectively those of the
    // last visible tab
    for ( i = 0; i < numTabs; i++ )
    {
        TabData = m_Tabs.Get( i );
        if ( TabData.visible )
        {
            *pBounds = TabData.bounds;
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the identifying number of the specified tab
//
//  Parameters:
//      tabIndex - zero based tab index
//
//  Returns:
//      DWORD id number.
//===============================================================================================//
DWORD TabWindow::GetTabIdentifier( size_t tabIndex ) const
{
    if ( tabIndex > m_Tabs.GetSize() )
    {
        throw BoundsException( L"tabIndex", __FUNCTION__ );
    }

    return m_Tabs.Get( tabIndex ).tabID;
}

//===============================================================================================//
//  Description:
//      Get the maximum allowed width of the tabs in pixels
//
//  Parameters:
//      none
//
//  Returns:
//      signed integer of maximum width, 0 if is there is no limit.
//===============================================================================================//
int TabWindow::GetMaxTabWidth() const
{
    return m_nMaxTabWidth;
}

//===============================================================================================//
//  Description:
//      Get the number of tabs
//
//  Parameters:
//      none
//
//  Returns:
//      size_t
//===============================================================================================//
size_t TabWindow::GetNumberTabs() const
{
    return m_Tabs.GetSize();
}

//===============================================================================================//
//  Description:
//      Get preferred tab strip height
//
//  Parameters:
//      None
//
//  Returns:
//      zero based selected tab number, -1 if none
//===============================================================================================//
int TabWindow::GetPreferredTabStripHeight() const
{
    int     height = 0;
    HDC     hDC;
    HFONT   hFont;
    HGDIOBJ hOldFont = nullptr;
    TEXTMETRIC tm;

    if ( m_hWindow == nullptr )
    {
        return MIN_TAB_HEIGHT;
    }

    hDC = GetDC( m_hWindow );
    if ( hDC == nullptr )
    {
        return MIN_TAB_HEIGHT;
    }

    try
    {
        hFont = m_Font.GetHandle();
    }
    catch ( const Exception& )
    {
        ReleaseDC( m_hWindow, hDC );
        throw;
    }

    if ( hFont )
    {
        hOldFont = SelectObject( hDC, hFont );
    }
    memset( &tm, 0, sizeof ( tm ) );
    GetTextMetrics( hDC, &tm );
    if ( hOldFont )
    {
        SelectObject( hDC, hOldFont );
    }
    ReleaseDC( m_hWindow, hDC );

    height  = tm.tmHeight + tm.tmExternalLeading;
    height += ( 2 * m_nPadding );
    height = PXSMaxInt( height, MIN_TAB_HEIGHT );
    height += INTERNAL_LEADING;
    height += BASELINE_HEIGHT;

    return height;
}

//===============================================================================================//
//  Description:
//      Get the selected tab index, zero based
//
//  Parameters:
//      None
//
//  Returns:
//      zero based selected tab number, -1 if none
//===============================================================================================//
size_t TabWindow::GetSelectedTabIndex() const
{
    return m_uSelectedTabIndex;
}

//===============================================================================================//
//  Description:
//      Get the name of the selected tab
//
//  Parameters:
//      pSelectedTabName - string object to receive the tab's name
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::GetSelectedTabName( String* pSelectedTabName ) const
{
    if ( pSelectedTabName == nullptr )
    {
        throw ParameterException( L"pSelectedTabName", __FUNCTION__ );
    }
    *pSelectedTabName = PXS_STRING_EMPTY;

    if ( m_uSelectedTabIndex < m_Tabs.GetSize() )
    {
       *pSelectedTabName = m_Tabs.Get( m_uSelectedTabIndex ).szName;
    }
}

//===============================================================================================//
//  Description:
//      Get the window of the selected tab
//
//  Parameters:
//      None
//
//  Returns:
//      pointer to the tab's window (aka page), NULL if no tab is selected
//===============================================================================================//
Window* TabWindow::GetSelectedTabWindow() const
{
    Window* pWindow = nullptr;

    if ( m_uSelectedTabIndex < m_Tabs.GetSize() )
    {
       pWindow = m_Tabs.Get( m_uSelectedTabIndex ).pWindow;
    }

    return pWindow;
}

//===============================================================================================//
//  Description:
//      Get the state of the specified tab
//
//  Parameters:
//      tabIndex - the tab's index
//
//  Returns:
//      DWORD defined constant
//===============================================================================================//
DWORD TabWindow::GetTabState( size_t tabIndex ) const
{
    if ( tabIndex > m_Tabs.GetSize() )
    {
        throw BoundsException( L"tabIndex", __FUNCTION__ );
    }

    return m_Tabs.Get( tabIndex ).state;
}

//===============================================================================================//
//  Description:
//      Get the window associated with a tab
//
//  Parameters:
//      tabIndex - the zero-based index of the tab
//
//  Returns:
//      pointer to the window, else null
//===============================================================================================//
Window* TabWindow::GetWindow( size_t tabIndex ) const
{
    if ( tabIndex > m_Tabs.GetSize() )
    {
        throw BoundsException( L"tabIndex", __FUNCTION__ );
    }

    return m_Tabs.Get( tabIndex ).pWindow;
}

//===============================================================================================//
//  Description:
//      Get the window associated with a tab
//
//  Parameters:
//      tabIdentifier - the tab's identifier
//
//  Returns:
//      pointer to the window, else null
//===============================================================================================//
Window* TabWindow::GetWindowFromIdentifier( DWORD tabIdentifier ) const
{
    size_t  i = 0;
    size_t  numTabs = m_Tabs.GetSize();
    Window* pWindow = nullptr;

    while( ( i < numTabs ) && ( pWindow == nullptr ) )
    {
        if ( tabIdentifier == GetTabIdentifier( i ) )
        {
            pWindow = m_Tabs.Get( i ).pWindow;
        }
        i++;
    }

    return pWindow;
}

//===============================================================================================//
//  Description:
//      Remove the specified tab
//
//  Parameters:
//      tabIndex = zero-based tab index
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::RemoveTab( size_t tabIndex )
{
    TYPE_TAB_DATA TabData;

    // Check bounds
    if ( tabIndex >= m_Tabs.GetSize() )
    {
        throw BoundsException( L"tabIndex", __FUNCTION__ );
    }
    TabData = m_Tabs.Get( tabIndex );

    if ( TabData.hBitmap )
    {
        DeleteObject( TabData.hBitmap );
    }
    m_Tabs.Remove( tabIndex );

    // Revert to the previous tab
    if ( m_Tabs.GetSize() )
    {
        if ( tabIndex == 0 )
        {
            SetSelectedTabIndex( 0 );
        }
        else
        {
            SetSelectedTabIndex( tabIndex - 1 );
        }
    }
    DoLayout();
}

//===============================================================================================//
//  Description:
//      Set the bitmaps used for the "close" image
//
//  Parameters:
//      closeID   - resource id of the default close bitmap
//      closeOnID - resource id of the highlighted close bitmap
//
//  The images should be small, no more that 16x16pixels. They are loaded
//  from the executable.
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::SetCloseBitmaps( WORD closeID, WORD closeOnID )
{
    HMODULE hModule = GetModuleHandle( nullptr );

    // Clean up
    if ( m_hBitmapClose )
    {
        DeleteObject( m_hBitmapClose );
        m_hBitmapClose = nullptr;
    }

    if ( m_hBitmapCloseOn )
    {
        DeleteObject( m_hBitmapCloseOn );
        m_hBitmapCloseOn = nullptr;
    }

    // Load
    if ( closeID )
    {
        m_hBitmapClose = LoadBitmap( hModule, MAKEINTRESOURCE( closeID ) );
        if ( m_hBitmapClose == nullptr )
        {
            throw SystemException( GetLastError(), L"LoadBitmap", __FUNCTION__ );
        }
    }

    if ( closeOnID )
    {
        m_hBitmapCloseOn = LoadBitmap( hModule, MAKEINTRESOURCE( closeOnID ) );
        if ( m_hBitmapCloseOn == nullptr )
        {
            throw SystemException( GetLastError(), L"LoadBitmap", __FUNCTION__ );
        }
    }
}

//===============================================================================================//
//  Description:
//      Set the maximum allowed width of the tabs in pixels
//
//  Parameters:
//      maxTabWidth - Maximum allowed width of the the tabs in pixels.
//                    Useful if the tab name is excessively long and
//                    do not want a very wide tab.
//                    Set to 0 if want no restriction on tab width limit
//                    in which case the tab will expand to fit its contents
//  Returns:
//      void
//===============================================================================================//
void TabWindow::SetMaxTabWidth( int maxTabWidth )
{
    // See if there is something to do
    if ( m_nMaxTabWidth != maxTabWidth )
    {
        m_nMaxTabWidth = PXSMaxInt( 0, maxTabWidth );
        DoLayout();
    }
}

//===============================================================================================//
//  Description:
//      Set the specified tab's name and tool tip
//
//  Parameters:
//      tabIndex   - the zero based tab index
//      Name    - the name/caption to show on the tab
//      ToolTip - tool tip text to show on mouse over
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::SetNameAndToolTip( size_t tabIndex, const String& Name, const String& ToolTip )
{
    TYPE_TAB_DATA TabData;

    // Check bounds
    if ( tabIndex >= m_Tabs.GetSize() )
    {
        throw BoundsException( L"tabIndex", __FUNCTION__ );
    }
    TabData = m_Tabs.Get( tabIndex );

    // The name, copy as much as fits
    if ( Name.GetLength() )
    {
        StringCchCopy( TabData.szName, ARRAYSIZE( TabData.szName ), Name.c_str() );
    }
    else
    {
        memset( TabData.szName, 0, sizeof ( TabData.szName ) );
    }

    // The tool tip, copy as much as fits
    if ( ToolTip.GetLength() )
    {
        StringCchCopy( TabData.szToolTip, ARRAYSIZE( TabData.szToolTip ), ToolTip.c_str() );
    }
    else
    {
        memset( TabData.szToolTip, 0, sizeof ( TabData.szToolTip ) );
    }
    m_Tabs.Set( tabIndex, TabData );
}

//===============================================================================================//
//  Description:
//      Set the specified tab as the selected one
//
//  Parameters:
//      setSelectedTabIndex - zero-based index of the tab
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::SetSelectedTabIndex( size_t setSelectedTabIndex )
{
    TYPE_TAB_DATA TabData;

    // Check bounds
    if ( setSelectedTabIndex >= m_Tabs.GetSize() )
    {
        throw BoundsException( L"setSelectedTabIndex", __FUNCTION__ );
    }
    m_uSelectedTabIndex = setSelectedTabIndex;
    TabData = m_Tabs.Get( m_uSelectedTabIndex );

    // Set the focus to the tab's page, to do so first do a layout
    // to ensure it is visible
    DoLayout();
    if ( TabData.pWindow && TabData.pWindow->GetHwnd() )
    {
        SetFocus( TabData.pWindow->GetHwnd() );
        TabData.pWindow->Repaint();
    }
    Repaint();
}

//===============================================================================================//
//  Description:
//      Set the tab's bitmap using the specified resource id
//
//  Parameters:
//      tabIndex   - the tab's array index
//      resourceID - the resoure id should be a 16x16 bitmap
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::SetTabBitmap( size_t tabIndex, WORD resourceID )
{
    TYPE_TAB_DATA TabData;

    if ( tabIndex >= m_Tabs.GetSize() )
    {
        throw BoundsException( L"tabIndex", __FUNCTION__ );
    }

    TabData = m_Tabs.Get( tabIndex );
    if ( TabData.hBitmap )
    {
        DeleteObject( TabData.hBitmap );
    }
    TabData.hBitmap = static_cast<HBITMAP>( LoadImage( GetModuleHandle(nullptr),
                                            MAKEINTRESOURCE( resourceID ),
                                            IMAGE_BITMAP, 16, 16, 0 ) );
    m_Tabs.Set( tabIndex, TabData );

    // Will avoid Repaint() as UpdateWindow() takes a lot of processing
    InvalidateRect( m_hWindow, &TabData.bounds, FALSE );
}

//===============================================================================================//
//  Description:
//      Set the high light colour used by the tab when it is selected
//
//  Parameters:
//      tabGradient1        - graident colour 1
//      tabGradient2        - graident colour 2
//      tabHiLiteBackground - the high light colour
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::SetTabColours( COLORREF tabGradient1,
                               COLORREF tabGradient2, COLORREF tabHiLiteBackground )
{
    m_crTabGradient1        = tabGradient1;
    m_crTabGradient2        = tabGradient2;
    m_crTabHiLiteBackground = tabHiLiteBackground;
}

//===============================================================================================//
//  Description:
//      Set the state of the specified tab
//
//  Parameters:
//      tabIndex - the tab's array index
//      state    - defined constant representing the state
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::SetTabState( size_t tabIndex, DWORD state )
{
    TYPE_TAB_DATA TabData;

    if ( tabIndex >= m_Tabs.GetSize() )
    {
        throw BoundsException( L"tabIndex", __FUNCTION__ );
    }
    TabData = m_Tabs.Get( tabIndex );
    TabData.state = state;
    m_Tabs.Set( tabIndex, TabData );
}

//===============================================================================================//
//  Description:
//      Set if tab strip height
//
//  Parameters:
//      tabStripHeight - the tab strip heigth, use zero for default height
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::SetTabStripHeight( int tabStripHeight )
{
    m_nTabStripHeight = tabStripHeight;
}

//===============================================================================================//
//  Description:
//      Set if the specified tab is visible
//
//  Parameters:
//      tabIndex - zero-based index of the tab
//      visible  - true if tab is visible, else false
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::SetTabVisible( size_t tabIndex, bool visible )
{
    size_t i = 0, selectedTab = 0;
    TYPE_TAB_DATA TabData;

    // Check bounds
    if ( tabIndex >= m_Tabs.GetSize() )
    {
        throw BoundsException( L"tabIndex", __FUNCTION__ );
    }
    TabData = m_Tabs.Get( tabIndex );

    // Set the tab
    if ( visible )
    {
        TabData.visible = TRUE;
        SetSelectedTabIndex( tabIndex );
    }
    else
    {
        // Hiding the specified tab, so select the one below to it
        TabData.visible = FALSE;
        for ( i = 0; i < tabIndex; i++ )
        {
            if ( m_Tabs.Get( tabIndex - 1 - i ).visible == TRUE )
            {
                selectedTab = tabIndex - 1 - i;
                break;
            }
        }
        SetSelectedTabIndex( selectedTab );
    }
    m_Tabs.Set( tabIndex, TabData );

    // Set the associated window
    if ( TabData.pWindow )
    {
        TabData.pWindow->SetVisible( visible );
    }
    DoLayout();
}

//===============================================================================================//
//  Description:
//      Set the tab width of all of the tabs
//
//  Parameters:
//      tabWidth - the tab width, use zero for default height
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::SetTabWidth( int tabWidth )
{
    if ( tabWidth < MIN_TAB_WIDTH )
    {
        tabWidth = MIN_TAB_WIDTH;
    }
    else if ( tabWidth > m_nMaxTabWidth )
    {
        tabWidth = m_nMaxTabWidth;
    }

    if ( tabWidth != m_nTabWidth )
    {
        m_nTabWidth = tabWidth;
        DoLayout();
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Handle the mouse left button down event
//
//  Parameters:
//      point - POINT where mouse was left clicked in the client area
//      keys  - which virtual keys are down
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::MouseLButtonDownEvent( const POINT& point, WPARAM /* keys */ )
{
    POINT    screenPos = point;
    size_t   tabIndex  = 0;
    WPARAM   wParam    = 0;
    ToolTip* pTip;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    // Forward event to tool tip
    pTip = ToolTip::GetInstance();
    if ( pTip )
    {
        // Need screen coordinates
        ClientToScreen( m_hWindow, &screenPos );
        pTip->MouseButtonDown( screenPos );
    }

    // Forward to listener
    if ( m_hWndAppMessageListener == nullptr )
    {
        return;     // No listener
    }

    tabIndex = MouseInTabIndex( point );
    if ( tabIndex >= m_Tabs.GetSize() )
    {
        return;     // Nothing to notify about
    }
    SetSelectedTabIndex( tabIndex );

    wParam = MAKEWPARAM( PXS_APP_MSG_ITEM_SELECTED, 0 );
    SendMessage( m_hWndAppMessageListener,
                 WM_COMMAND, wParam, (LPARAM)m_hWindow );
}

//===============================================================================================//
//  Description:
//      Handle the mouse left button up event
//
//  Parameters:
//      point - POINT where mouse was left clicked in the client area
//      keys  - which virtual keys are down
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::MouseLButtonUpEvent( const POINT& point )
{
    RECT     bounds    = { 0, 0, 0, 0 };
    size_t   tabIndex  = 0;
    WPARAM   wParam    = 0;
    BITMAP   bmClose;
    TYPE_TAB_DATA TabData;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    if ( m_hWndAppMessageListener == nullptr )
    {
        return;     // No listener
    }

    // Determine if clicked on the close bitmap, if so will close it provided
    // there is more than one tab
    if ( m_Tabs.GetSize() <= 1 )
    {
        return;
    }
    tabIndex = MouseInTabIndex( point );
    if ( tabIndex >= m_Tabs.GetSize() )
    {
        return;     // Nothing to notify about
    }

    memset( &bounds   , 0, sizeof ( bounds ) );
    memset( &bmClose, 0, sizeof ( bmClose ) );
    TabData = m_Tabs.Get( tabIndex );
    if ( ( m_hBitmapClose == nullptr ) || ( TabData.allowClose == false ) )
    {
        return;     // Not showing a bitmap
    }
    GetObject( m_hBitmapClose, sizeof ( bmClose ), &bmClose );

    // Set bounds
    if ( IsRightToLeftReading() )
    {
        bounds.left = TabData.bounds.left + HORIZ_TAB_GAP;
    }
    else
    {
        bounds.left = TabData.bounds.right - bmClose.bmWidth - HORIZ_TAB_GAP;
    }
    bounds.right  = bounds.left + bmClose.bmWidth;
    bounds.top    = TabData.bounds.top + ( ( TabData.bounds.bottom
                                             - TabData.bounds.top
                                             - bmClose.bmHeight ) / 2 );
    bounds.bottom = bounds.top  + bmClose.bmHeight;

    if ( PtInRect( &bounds, point ) )
    {
        // Send a close message
        wParam = MAKEWPARAM( PXS_APP_MSG_REMOVE_ITEM, 0 );
        SendMessage( m_hWndAppMessageListener, WM_COMMAND, wParam, (LPARAM)m_hWindow);
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
//  Returns:
//      void
//===============================================================================================//
void TabWindow::MouseMoveEvent( const POINT& point, WPARAM /* keys */ )
{
    size_t   tabIndex = 0;
    String   TipText;
    ToolTip* pTip = nullptr;
    TYPE_TAB_DATA TabData;

    // Class scope data checks
    if ( m_hWindow == nullptr )
    {
        return;
    }
    SetCursor( LoadCursor( nullptr, IDC_ARROW ) );

    // If this window is not contained in the foreground window
    // then ignore rest of processing
    if ( IsChild( GetForegroundWindow(), m_hWindow )  == 0 )
    {
        return;
    }

    // Determine if the mouse is inside a tab, if so show its tool tip
    tabIndex = MouseInTabIndex( point );
    if ( tabIndex < m_Tabs.GetSize() )
    {
        SetTimer( m_hWindow, m_uTimerID, 100, nullptr );

        pTip = ToolTip::GetInstance();
        if ( pTip )
        {
            // Redraw, in case the tab has a close button
            Repaint();
            TabData = m_Tabs.Get( tabIndex );
            if ( tabIndex != m_uMouseInTabIndex )
            {
                RectangleToScreen( &TabData.bounds );
                TipText = TabData.szToolTip;
                pTip->Show( TipText, TabData.bounds );
            }
        }
    }

    // Store for the next mouse event
    if ( tabIndex != m_uMouseInTabIndex )
    {
        m_uMouseInTabIndex = tabIndex;
        Repaint();
    }
}

//===============================================================================================//
//  Description:
//      Handle the WM_RBUTTONDOWN event
//
//  Parameters:
//      point - POINT where mouse was right clicked in the client area
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::MouseRButtonDownEvent( const POINT& point )
{
    POINT screenPos = point;
    ToolTip* pTip   = nullptr;

    // Forward event to tool tip
    pTip = ToolTip::GetInstance();
    if ( pTip )
    {
        // Need screen coordinates
        ClientToScreen( m_hWindow, &screenPos );
        pTip->MouseButtonDown( screenPos );
    }
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
void TabWindow::NotifyEvent( const NMHDR* pNmhdr )
{
    // Forward to the event listener if there is one
    if ( m_hWndAppMessageListener )
    {
        SendMessage( m_hWndAppMessageListener, WM_NOTIFY, 0, (LPARAM)pNmhdr );
    }
}

//===============================================================================================//
//  Description:
//      Paint event handler
//
//  Parameters:
//      hdc - Handle to device context
//
//  Remarks:
//      This method and DoLayout need to be kept in sync.
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::PaintEvent( HDC hdc )
{
    SIZE   size    = { 0, 0 };
    RECT   bounds  = { 0, 0, 0, 0 };
    size_t numTabs = 0, i = 0;
    TYPE_TAB_DATA TabData;
    StaticControl Static;

    if ( hdc == nullptr )
    {
        return;
    }
    DrawBackground( hdc );

    numTabs = m_Tabs.GetSize();
    for ( i = 0; i < numTabs; i++ )
    {
        TabData = m_Tabs.Get( i );
        if ( TabData.visible )
        {
            DrawTabShape( hdc, i );
            DrawTabBitmapAndText( hdc, TabData );
            if ( TabData.allowClose && ( numTabs > 1 ) )
            {
                DrawTabCloseBitmap( hdc, TabData.bounds );
            }
        }
    }

    // Draw the base line
    GetTabBounds( &bounds );
    GetClientSize( &size );
    bounds.left   = 0;
    bounds.top    = bounds.bottom - BASELINE_HEIGHT;
    bounds.right  = size.cx;
    Static.SetBounds( bounds );
    Static.SetShape( PXS_SHAPE_RECTANGLE );
    Static.SetShapeColour( m_crForeground );
    Static.Draw( hdc );

    // Draw the border for the tab strip
    if ( ( m_uBorderStyle != PXS_SHAPE_NONE ) && m_nTabStripHeight )
    {
        bounds.top    = 0;
        bounds.bottom = m_nTabStripHeight - 1;
        Static.Reset();
        Static.SetBounds( bounds );
        Static.SetShape( m_uBorderStyle );
        Static.SetShapeColour( GetSysColor( COLOR_3DDKSHADOW ) );
        Static.Draw( hdc );
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
void TabWindow::TimerEvent( UINT_PTR timerID )
{
    POINT cursorPos  = { 0, 0 };

    GetCursorPos( &cursorPos );
    ScreenToClient( m_hWindow, &cursorPos );
    m_uMouseInTabIndex = MouseInTabIndex( cursorPos );
    if ( m_uMouseInTabIndex == PXS_MINUS_ONE )
    {
        KillTimer( m_hWindow, timerID );
        Repaint();
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

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
void TabWindow::DrawTabCloseBitmap( HDC hDC, const RECT& tabBounds )
{
    RECT    bounds = tabBounds;
    POINT   cursor = { 0, 0 };
    BITMAP  bmClose;
    StaticControl Static;

    if ( ( hDC == nullptr ) || ( m_hBitmapClose == nullptr ) )
    {
        return;     // Nothing to do
    }

    // Will assume the default and "on" are the same size
    memset( &bmClose, 0, sizeof ( bmClose ) );
    if ( GetObject( m_hBitmapClose, sizeof ( bmClose ), &bmClose ) == 0 )
    {
        return;
    }

    // Set bounds
    if ( IsRightToLeftReading() )
    {
        bounds.left += HORIZ_TAB_GAP;
    }
    else
    {
        bounds.left = bounds.right - bmClose.bmWidth - HORIZ_TAB_GAP;
    }
    bounds.right  = bounds.left + bmClose.bmWidth;
    bounds.top    = bounds.top  + (bounds.bottom-bounds.top-bmClose.bmHeight)/2;
    bounds.bottom = bounds.top  + bmClose.bmHeight;

    // Highlighted image of selected tab on mouse over
    if ( ( GetCursorPos( &cursor )  ) &&
         ( ScreenToClient( m_hWindow, &cursor ) ) &&
         ( PtInRect( &bounds, cursor ) ) )
    {
        Static.SetBitmap( m_hBitmapCloseOn );
    }
    else
    {
        Static.SetBitmap( m_hBitmapClose );
    }
    Static.SetBounds( bounds );
    Static.SetPadding( 0 );
    Static.Draw( hDC );
}

//===============================================================================================//
//  Description:
//      Draw the specified tab
//
//  Parameters:
//      hDC      - handle to the device context
//      tabIndex - zero based index
//
//  Returns:
//     void
//===============================================================================================//
void TabWindow::DrawTabShape( HDC hDC, size_t tabIndex )
{
    bool  rtlReading;
    RECT  bounds;
    DWORD tabShape = PXS_SHAPE_TAB;
    TYPE_TAB_DATA TabData;
    StaticControl Static;

    if ( hDC == nullptr )
    {
        return;
    }

    if ( tabIndex > m_Tabs.GetSize() )
    {
        throw BoundsException( L"tabIndex", __FUNCTION__ );
    }

    TabData = m_Tabs.Get( tabIndex );
    bounds  = TabData.bounds;
    if ( tabIndex == m_uSelectedTabIndex )
    {
        Static.SetBackground( m_crTabHiLiteBackground );
    }
    else
    {
        Static.SetBackgroundGradient( m_crTabGradient1, m_crTabGradient2, true );
    }

    rtlReading = IsRightToLeftReading();
    if ( tabIndex == m_uMouseInTabIndex )
    {
        bounds.bottom -= BASELINE_HEIGHT;
        if ( rtlReading )
        {
            tabShape = PXS_SHAPE_TAB_DOTTED_RTL;
        }
        else
        {
            tabShape = PXS_SHAPE_TAB_DOTTED;
        }
    }
    else if ( rtlReading )
    {
        tabShape = PXS_SHAPE_TAB_RTL;
    }
    Static.SetShape( tabShape );
    Static.SetShapeColour( m_crForeground );
    Static.SetBounds( bounds );
    Static.Draw( hDC );
}

//===============================================================================================//
//  Description:
//      Draw the tab's text
//
//  Parameters:
//      hDC       - handle to the device context
//      pszText   - the text
//      tabBounds - the tab's bounds
//
//  Remarks:
//      Always draw the close button last
//      Layout: gap | bitmap | gap | text | gap | close image | gap
//
//  Returns:
//      Zero-based tab number the pointer is in. If the pointer is not in a
//      tab returns PXS_MINUS_ONE
//
//===============================================================================================//
void TabWindow::DrawTabBitmapAndText( HDC hDC, const TYPE_TAB_DATA& TabData )
{
    bool  rightToLeft = IsRightToLeftReading();
    int   tabHeight, requiredWidth, oldBkMode;
    HFONT hFont;
    RECT  bounds     = { 0, 0, 0, 0 };
    UINT  textFormat = DT_SINGLELINE | DT_VCENTER | DT_END_ELLIPSIS;
    BITMAP  bmBitmap, bmClose;
    HGDIOBJ hOldFont = nullptr;
    StaticControl Static;

    if ( hDC == nullptr )
    {
        return;
    }
    tabHeight = TabData.bounds.bottom - TabData.bounds.top;

    // Bitmaps
    memset( &bmClose, 0, sizeof ( bmClose ) );
    if ( m_hBitmapClose )
    {
        GetObject( m_hBitmapClose, sizeof ( bmClose ), &bmClose );
    }

    memset( &bmBitmap, 0, sizeof ( bmBitmap ) );
    if ( TabData.hBitmap )
    {
        GetObject( TabData.hBitmap, sizeof ( bmBitmap ), &bmBitmap );
    }

    // Determine if there is enough room for the bitmaps
    requiredWidth = HORIZ_TAB_GAP + bmBitmap.bmWidth;
    if ( TabData.allowClose )
    {
        requiredWidth += bmClose.bmWidth + HORIZ_TAB_GAP;
    }

    if ( requiredWidth > ( TabData.bounds.right - TabData.bounds.left ) )
    {
        return;
    }

    // Draw bitmap
    if ( bmBitmap.bmWidth )
    {
        if ( rightToLeft )
        {
            bounds.left = TabData.bounds.right - bmBitmap.bmWidth - HORIZ_TAB_GAP;
        }
        else
        {
            bounds.left = TabData.bounds.left + HORIZ_TAB_GAP;
        }
        bounds.right = bounds.left + bmBitmap.bmWidth;
        bounds.top    = TabData.bounds.top + (tabHeight - bmBitmap.bmHeight)/2;
        bounds.bottom = bounds.top + bmBitmap.bmHeight;
        Static.SetBitmap( TabData.hBitmap );
        Static.SetBounds( bounds );
        Static.SetPadding( 0 );
        Static.Draw( hDC );
    }

    // Text area
    bounds = TabData.bounds;
    if ( rightToLeft )
    {
        textFormat  |= ( DT_RIGHT | DT_RTLREADING );
        bounds.left  = TabData.bounds.left
                       + HORIZ_TAB_GAP + bmClose.bmWidth + HORIZ_TAB_GAP;
        bounds.right = TabData.bounds.right
                       - HORIZ_TAB_GAP - bmBitmap.bmWidth - HORIZ_TAB_GAP;
    }
    else
    {
        bounds.left  = TabData.bounds.left
                       + HORIZ_TAB_GAP + bmBitmap.bmWidth + HORIZ_TAB_GAP;
        bounds.right = TabData.bounds.right
                       - HORIZ_TAB_GAP - bmClose.bmWidth - HORIZ_TAB_GAP;
    }

    if ( bounds.right <= bounds.left )
    {
        return;     // No room to draw the text
    }
    hFont = m_Font.GetHandle();
    if ( hFont )
    {
        hOldFont = SelectObject( hDC, hFont );
    }
    oldBkMode = SetBkMode( hDC, TRANSPARENT );
    DrawText( hDC, TabData.szName, lstrlen( TabData.szName ), &bounds, textFormat );

    // Reset DC
    if ( oldBkMode ) SetBkMode( hDC, oldBkMode );
    if ( hOldFont  ) SelectObject( hDC, hOldFont );
}

//===============================================================================================//
//  Description:
//      Evaluate the preferred with of a tab in pixels
//
//  Parameters:
//      hDC            - the device context onto which the tab is to be drawn
//      tabHeight      - the tab's height
//      pszTabText     - the text to display on the tab
//      showCloseImage - indicates if want to display the close button image
//
//  Remarks:
//      Layout is (LTR): gap | bitmap | gap | text | gap | close image | gap
//
//  Returns:
//      width of the tab in pixels
//===============================================================================================//
int TabWindow::EvaluateTabWidth( HDC hDC, int tabHeight, const TYPE_TAB_DATA& TabData )
{
    int    tabWidth = tabHeight;        // 45 degree slope
    SIZE   sizeText = { 0, 0 };
    BITMAP bmImage;

    // If a width was specified, use that
    if ( m_nTabWidth > 0 )
    {
        return PXSMinInt( m_nTabWidth, m_nMaxTabWidth );
    }

    // Bitmap image
    tabWidth = HORIZ_TAB_GAP;
    memset( &bmImage, 0, sizeof ( bmImage ) );
    if ( TabData.hBitmap )
    {
        memset( &bmImage, 0, sizeof ( bmImage ) );
        if ( GetObject( TabData.hBitmap, sizeof ( bmImage ), &bmImage ) )
        {
            tabWidth += bmImage.bmWidth;
        }
    }
    tabWidth += HORIZ_TAB_GAP;

    // Add the width of the text
    if ( hDC && TabData.szName[ 0 ] )
    {
        if ( GetTextExtentPoint32( hDC,
                                   TabData.szName,
                                   lstrlen( TabData.szName ), &sizeText ) )
        {
            tabWidth += sizeText.cx;
        }
    }
    tabWidth += HORIZ_TAB_GAP;

    // Add space for the close image even if do not show it
    if ( m_hBitmapClose )
    {
        memset( &bmImage, 0, sizeof ( bmImage ) );
        if ( GetObject( m_hBitmapClose, sizeof ( bmImage ), &bmImage ))
        {
            tabWidth += bmImage.bmWidth;
        }
    }
    tabWidth += HORIZ_TAB_GAP;

    // Limit the tab width. A non-positive value for maximum implies no limit
    if ( tabWidth < MIN_TAB_WIDTH )
    {
        tabWidth = MIN_TAB_WIDTH;
    }

    if ( m_nMaxTabWidth > 0 )
    {
        tabWidth = PXSMinInt( tabWidth, m_nMaxTabWidth );
    }

    return tabWidth;
}

//===============================================================================================//
//  Description:
//      Layout the pages belonging to the tabs
//
//  Parameters:
//      bounds - the bounds for the pages
//
//  Returns:
//      void
//===============================================================================================//
void TabWindow::LayoutPages( const RECT& bounds )
{
    HWND   hWndPage = nullptr;
    size_t i = 0;
    size_t numTabs = m_Tabs.GetSize();
    TYPE_TAB_DATA TabData;

    for ( i = 0; i < numTabs; i++ )
    {
        TabData = m_Tabs.Get( i );
        if ( TabData.pWindow )
        {
            hWndPage = TabData.pWindow->GetHwnd();
            if ( hWndPage )
            {
                // Bring the page of the selected to the top
                if ( i == m_uSelectedTabIndex )
                {
                    SetWindowPos( hWndPage,
                                  nullptr,      // = HWND_TOP,
                                  bounds.left,
                                  bounds.top,
                                  bounds.right  - bounds.left,
                                  bounds.bottom - bounds.top, SWP_SHOWWINDOW );
                }
                else
                {
                    ShowWindow( hWndPage, SW_HIDE );
                    MoveWindow( hWndPage,
                                bounds.left,
                                bounds.top,
                                bounds.right  - bounds.left,
                                bounds.bottom - bounds.top, FALSE );
                }
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Get the zero-based tab index that the cursor is in.
//
//  Parameters:
//      point - POINT structure containing mouse coordinates in the client area
//
//  Returns:
//      Zero-based tab number the pointer is in. If the pointer is not in a
//      tab returns PXS_MINUS_ONE
//===============================================================================================//
size_t TabWindow::MouseInTabIndex( const POINT& point )
{
    RECT   bounds = { 0, 0, 0, 0 };
    size_t i = 0, numTabs = 0, tabIndex = PXS_MINUS_ONE;

    // Class scope data checks
    if ( m_hWindow == nullptr )
    {
        return PXS_MINUS_ONE;
    }

    // Get tab's bounds and see if cursor is inside
    numTabs = m_Tabs.GetSize();
    for ( i = 0; i < numTabs; i++ )
    {
        bounds = m_Tabs.Get( i ).bounds;
        if ( PtInRect( &bounds, point ) )
        {
            tabIndex = i;
            break;
        }
    }

    return tabIndex;
}
