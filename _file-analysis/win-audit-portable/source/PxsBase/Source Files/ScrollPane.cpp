///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Scroll Pane Class Implementation
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
#include "PxsBase/Header Files/ScrollPane.h"

// 2. C System Files
#include <stdint.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/ParameterException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
ScrollPane::ScrollPane()
           :m_sMouseWheelDelta( 0 ),
            m_nScreenLineHeight( PXS_DEFAULT_SCREEN_LINE_HEIGHT )
{
    // Window creation
    m_CreateStruct.style = WS_CHILD | WS_VISIBLE | WS_HSCROLL | WS_VSCROLL;
}

// Copy constructor - not allowed so no implementation

// Destructor
ScrollPane::~ScrollPane()
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
//      Get the number of lines visible in the client area
//
//  Parameters:
//      none
//
//  Returns:
//      number of lines
//===============================================================================================//
int ScrollPane::GetNumClientLines()
{
    SIZE clientSize = { 0, 0 };

    GetClientSize( &clientSize );
    if ( ( clientSize.cy > 0 ) && ( m_nScreenLineHeight > 0 ) )
    {
        return ( clientSize.cy / m_nScreenLineHeight );
    }
    return 0;
}

//===============================================================================================//
//  Description:
//      Get the scroll port's position        .
//
//  Parameters:
//      pScrollPosition - a POINT to receive the position
//
//  Returns:
//      true if got the position else false.
//===============================================================================================//
bool ScrollPane::GetScrollPosition( POINT* pScrollPosition )
{
    bool bOK = true;
    SCROLLINFO si;

    if ( pScrollPosition == nullptr )
    {
        throw ParameterException( L"pScrollPosition", __FUNCTION__ );
    }
    pScrollPosition->x = 0;
    pScrollPosition->y = 0;

    if ( m_hWindow )
    {
        // X position
        memset( &si, 0, sizeof ( si ) );
        si.cbSize = sizeof ( si );
        si.fMask  = SIF_ALL;
        if ( GetScrollInfo( m_hWindow, SB_HORZ, &si ) )
        {
            pScrollPosition->x = si.nPos;
        }
        else
        {
            bOK = false;
        }

        // Y position
        memset( &si, 0, sizeof ( si ) );
        si.cbSize = sizeof ( si );
        si.fMask  = SIF_ALL;
        if ( GetScrollInfo( m_hWindow, SB_VERT, &si ) )
        {
            pScrollPosition->y = si.nPos;
        }
        else
        {
            bOK = false;
        }
    }

    return bOK;
}

//===============================================================================================//
//  Description:
//      Set the scroll panes's position.
//
//  Parameters:
//      scrollPosition - a POINT containing the new position
//
//  Returns:
//      true if set position, else false
//===============================================================================================//
bool ScrollPane::SetScrollPosition( const POINT& scrollPosition )
{
    bool bOK  = false;
    int  xPos = 0, yPos = 0;
    SCROLLINFO si;

    if ( m_hWindow == nullptr )
    {
        return false;
    }

    // X position
    memset( &si, 0, sizeof ( si ) );
    si.cbSize = sizeof ( si );
    si.fMask  = SIF_POS;
    si.nPos   = scrollPosition.x;
    xPos = SetScrollInfo( m_hWindow, SB_HORZ, &si, TRUE );

    // Y position
    memset( &si, 0, sizeof ( si ) );
    si.cbSize = sizeof ( si );
    si.fMask  = SIF_POS;
    si.nPos   = scrollPosition.y;
    yPos = SetScrollInfo( m_hWindow, SB_VERT, &si, TRUE );

    Repaint();

    // Verify scroll bar was set
    if ( ( xPos == scrollPosition.x ) && ( yPos == scrollPosition.y ) )
    {
        bOK = true;
    }

    return bOK;
}

//===============================================================================================//
//  Description:
//      Update the SCROLLINFO structure
//
//  Parameters:
//        size - SIZE structure of the scrollable area
//
//  Returns:
//      void
//===============================================================================================//
void ScrollPane::UpdateScrollBarsInfo( const SIZE& size )
{
    int  width      = 0;
    UINT pageNumber = 0;
    RECT clientRect = { 0, 0, 0, 0 };
    SCROLLINFO si;
    TEXTMETRIC tm;

    if ( m_hWindow == nullptr )
    {
        return;
    }
    GetClientRect( m_hWindow, &clientRect );

    ////////////////////////////////////////////////
    // Horizontal scroll bar

    memset( &si, 0, sizeof ( si ) );
    si.cbSize = sizeof ( si );
    si.fMask  = SIF_ALL;
    GetScrollInfo( m_hWindow, SB_HORZ, &si );

    si.nMin = 0;
    width   = clientRect.right - clientRect.left;
    if ( width > 0 )
    {
        memset( &tm, 0, sizeof ( tm ) );
        PXSGetStockFontTextMetrics( nullptr, DEFAULT_GUI_FONT, &tm );
        if ( tm.tmAveCharWidth > 0 )
        {
            pageNumber = PXSCastLongToUInt32( width / tm.tmAveCharWidth );
        }
    }

    // Only apply if there is a change to avoid raising a WM_SIZE message
    if ( ( si.nMax != size.cx ) || ( si.nPage != pageNumber ) )
    {
        si.nMax  = size.cx;
        si.nPage = pageNumber;
        SetScrollInfo( m_hWindow, SB_HORZ, &si, TRUE );
    }

    ////////////////////////////////////////////////
    // Vertical scroll bar

    memset( &si, 0, sizeof ( si ) );
    si.cbSize = sizeof ( si );
    si.fMask  = SIF_ALL;
    GetScrollInfo( m_hWindow, SB_VERT, &si );

    si.nMin    = 0;
    pageNumber = 0;
    if ( clientRect.bottom > clientRect.top )
    {
        pageNumber = PXSCastLongToUInt32( clientRect.bottom - clientRect.top );
    }

    if ( ( si.nMax != size.cy ) || ( si.nPage != pageNumber ) )
    {
        si.nMax  = size.cy;
        si.nPage = pageNumber;
        SetScrollInfo( m_hWindow, SB_VERT, &si, TRUE );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Handle WM_KEYDOWN  event for this window.
//
//  Parameters:
//      virtKey - signed integer of virtual key code
//
//  Returns:
//      LRESULT, 0 if it the message was processed else 1.
//===============================================================================================//
LRESULT ScrollPane::KeyDownEvent( WPARAM virtKey )
{
    bool vertical   = false;
    int  scrollCode = -1;  // Scroll codes are 0-8 defined in WinUser.h

    // Interpret the key pressed
    switch ( virtKey )
    {
        default:
            break;

        ///////////////////////////////////////////////////
        // Key strokes that correspond to vertical movement

        case VK_PRIOR:              // page up key
            vertical   = true;
            scrollCode = SB_PAGEUP;
        break;

        case VK_NEXT:               // page down key
            vertical   = true;
            scrollCode = SB_PAGEDOWN;
        break;

        case VK_END:                // end key
            vertical   = true;
            scrollCode = SB_BOTTOM;
        break;

        case VK_HOME:               // home key
            vertical   = true;
            scrollCode = SB_TOP;
        break;

        case VK_UP:                 // up arrow key
            vertical   = true;
            scrollCode = SB_LINEUP;
        break;

        case VK_DOWN:               // down arrow key
            vertical   = true;
            scrollCode = SB_LINEDOWN;
        break;

        /////////////////////////////////////////////////////
        // Key strokes that correspond to horizontal movement

        case VK_LEFT:               // left arrow key
            scrollCode = SB_LINEUP;
        break;

        case VK_RIGHT:              // right arrow key
            scrollCode = SB_LINEDOWN;
        break;
    }

    // Call the event handler directly
    return ScrollEvent( vertical, scrollCode );
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
//  Remarks::
//        Normally an wheel event is a delta of 120, but can
//        be more fine grained, so keep a total of the deltas until
//        get to 120 then scroll the window by one line
//
//        MSDN: The MouseWheel has 18 'detents'. As the wheel
//        is rolled and a detent is encountered, the OS will
//        send a WM_MOUSEWHEEL message with the HIWORD of
//        wParam set to a value of +/- 120.
//        '+' if the wheel is being rolled forward (away from
//        the user), '-' if the wheel is being rolled backwards
//        (towards the user).
//
//        Rolled towards user   == scroll downwards = -120
//        Rolled away from user == scroll upwards   = +120
//
//  Returns:
//      void
//===============================================================================================//
void ScrollPane::MouseWheelEvent( WORD /* keys */,
                                  SHORT delta, const POINT& point )
{
    UINT  ParametersInfo   = 0;
    long  wheelScrollLines = 3;     // Default value is 3
    POINT scrollPosition   = { 0, 0 };

    if ( ( ( m_sMouseWheelDelta + delta ) > SHRT_MAX ) ||
         ( ( m_sMouseWheelDelta + delta ) < SHRT_MIN )  )
    {
        return;     // Out of range
    }
    m_sMouseWheelDelta = (SHORT)(m_sMouseWheelDelta + delta);

    if ( SystemParametersInfo( SPI_GETWHEELSCROLLLINES, 0, &ParametersInfo, 0 ) )
    {
        wheelScrollLines = (long)ParametersInfo;
    }

    // If have rolled at least a WHEEL_DELTA, then scroll the window.
    if ( abs( m_sMouseWheelDelta ) >= WHEEL_DELTA )
    {
        GetScrollPosition( &scrollPosition );
        if ( m_sMouseWheelDelta > 0 )
        {
            scrollPosition.y -= ( wheelScrollLines * m_nScreenLineHeight );
        }
        else
        {
            scrollPosition.y += ( wheelScrollLines * m_nScreenLineHeight );
        }
        SetScrollPosition( scrollPosition );

        // Throw away any amount greater than WHEEL_DELTA
        m_sMouseWheelDelta = 0;

        // Raise a mouse move event, as the cursor shape may be
        // determined by what is underneath it, e.g. a hyper-link
        SendMessage( m_hWindow, WM_MOUSEMOVE, 0, MAKELPARAM( point.x, point.y ) );
        InvalidateRect( m_hWindow, nullptr, FALSE );
    }
}

//===============================================================================================//
//  Description:
//  Handle WM_HSCROLL and WM_VSCROLL  events.
//
//  Parameters:
//      vertical    - if true this is a vertical scroll event, else horizontal
//      scrollCode  - a code that indicates the type of scroll movement
//
//  Returns:
//      LRESULT, 0 if it the message was processed else 1.
//===============================================================================================//
LRESULT ScrollPane::ScrollEvent( bool vertical, int scrollCode )
{
    int   fnBar  = SB_HORZ, originalPosition = 0;
    POINT cursor = { 0, 0 };
    SCROLLINFO si;

    // Ensure window was created
    if ( m_hWindow == nullptr )
    {
        return 1;   // Not handles
    }

    // Get the scrollbar's current settings
    memset( &si, 0, sizeof ( si ) );
    si.cbSize = sizeof ( si );
    si.fMask  = SIF_ALL;
    if ( vertical )
    {
        fnBar = SB_VERT;
    }
    if ( GetScrollInfo( m_hWindow, fnBar, &si ) == 0 )
    {
        return 1;       // not handled
    }
    originalPosition = si.nPos;

    // Codes are symmetrical for most movements, e.g. SB_LINEUP == SB_LINELEFT
    switch ( scrollCode )
    {
        default:
            break;

        case SB_LINELEFT:   // = case SB_LINEUP
            si.nPos -= m_nScreenLineHeight;
            break;

        case SB_LINERIGHT:  // = SB_LINEDOWN
            si.nPos += m_nScreenLineHeight;
            break;

        case SB_PAGELEFT:   // = SB_PAGEUP
            si.nPos -= PXSCastUInt32ToInt32( si.nPage );
            break;

        case SB_PAGERIGHT:  // = SB_PAGEDOWN
            si.nPos += PXSCastUInt32ToInt32( si.nPage );
            break;

        // Fall through
        case SB_THUMBPOSITION:
        case SB_THUMBTRACK:
            si.nPos = si.nTrackPos;
            break;

        case SB_LEFT:       // = SB_TOP
            si.nPos = si.nMin;
            break;

        case SB_RIGHT:      // = SB_BOTTOM
            si.nPos = si.nMax;
            break;

        case SB_ENDSCROLL:

            // Raise a mouse move event as the cursor shape may be
            // determined by what is underneath it, e.g. a hyper-link
            GetCursorPos( &cursor );
            SendMessage( m_hWindow, WM_MOUSEMOVE, 0, MAKELPARAM( cursor.x, cursor.y ) );
            break;
    }

    // Only need to update the scroll port if the position has changed
    if ( si.nPos != originalPosition )
    {
        SetScrollInfo( m_hWindow, fnBar, &si, TRUE );
        DoLayout();
        InvalidateRect( m_hWindow, nullptr, FALSE );
    }

    return 0;   // = event was handled
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
