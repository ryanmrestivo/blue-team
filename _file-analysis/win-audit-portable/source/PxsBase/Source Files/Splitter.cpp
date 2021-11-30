///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Splitter Class Implementation
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
#include "PxsBase/Header Files/Splitter.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/NullException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StaticControl.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
Splitter::Splitter()
         :DIVIDER_THICKNESS( 9 ),
          m_bPanelsAttached( false ),
          m_bMouseWasPressed( false ),
          m_hWndOne( nullptr ),
          m_hWndTwo( nullptr ),
          m_SplitterLocation(),            // Non-op
          m_InitialCursorLocation()        // Non-op

{
    memset( &m_SplitterLocation     , 0, sizeof ( m_SplitterLocation      ) );
    memset( &m_InitialCursorLocation, 0, sizeof ( m_InitialCursorLocation ) );

    // Properties
    m_crBackground = GetSysColor( COLOR_BTNFACE );
    try
    {
        SetForeground( PXS_COLOUR_NAVY );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy constructor - not allowed so no implementation

// Destructor
Splitter::~Splitter()
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
//      Attach two windows to this splitter window and do initial layout.
//
//  Parameters:
//      pOne - window one (left)
//      pTwo - window two (right)
//
//  Remarks::
//      The splitter window must have been created already.
//
//  Returns:
//      void
//===============================================================================================//
void Splitter::AttachWindows( const Window* pOne, const Window* pTwo )
{
    int  offset = 100;  // Initial layout with the value of 100
    RECT rect   = { 0, 0, 0, 0 };

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    if ( ( pOne == nullptr ) || ( pTwo == nullptr ) )
    {
        throw ParameterException( L"pOne/pTwo", __FUNCTION__ );
    }

    m_hWndOne = pOne->GetHwnd();
    m_hWndTwo = pTwo->GetHwnd();
    if ( ( m_hWndOne == nullptr ) || (m_hWndTwo == nullptr ) )
    {
        throw NullException( L"m_hWndOne/m_hWndTwo", __FUNCTION__ );
    }
    m_bPanelsAttached = true;

    if ( GetClientRect( GetParent( m_hWindow ), &rect ) )
    {
        LONG W = rect.right  - rect.left;
        LONG H = rect.bottom - rect.top;

        MoveWindow( m_hWndOne, 0, 0, offset, H, TRUE );
        MoveWindow( m_hWindow, offset, 0, DIVIDER_THICKNESS, H, TRUE );
        MoveWindow( m_hWndTwo,
                    offset + DIVIDER_THICKNESS, 0, W - offset - DIVIDER_THICKNESS, H, TRUE );
    }
}

//===============================================================================================//
//  Description:
//      Layout the splitter window and the two attached panels
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Splitter::DoLayout()
{
    HWND hWndParent = GetParent( m_hWindow );
    RECT parentWindowRect = { 0, 0, 0, 0 };
    RECT parentClientRect = { 0, 0, 0, 0 };
    RECT thisWindowRect   = { 0, 0, 0, 0 };

    // Ensure panels are attached
    if ( m_bPanelsAttached == false ) return;

    if ( GetWindowRect( hWndParent, &parentWindowRect ) &&
         GetClientRect( hWndParent, &parentClientRect ) &&
         GetWindowRect( m_hWindow , &thisWindowRect   )  )
    {
        LONG W    = parentClientRect.right  - parentClientRect.left;
        LONG H    = parentClientRect.bottom - parentClientRect.top;
        LONG xPos = thisWindowRect.left - parentWindowRect.left;

        MoveWindow( m_hWndOne, 0, 0, xPos, H, FALSE );
        MoveWindow( m_hWindow, xPos, 0, DIVIDER_THICKNESS, H, FALSE );
        MoveWindow( m_hWndTwo,
                    xPos + DIVIDER_THICKNESS, 0, W - (xPos + DIVIDER_THICKNESS), H, FALSE );
    }
    RedrawWindow(m_hWndOne, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN);
    RedrawWindow(m_hWindow, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN);
    RedrawWindow(m_hWndTwo, nullptr, nullptr, RDW_INVALIDATE | RDW_ALLCHILDREN);
}

//===============================================================================================//
//  Description:
//      Get the offset of the divider window
//
//  Parameters:
//      None
//
//  Returns:
//      int of offset from left or top
//===============================================================================================//
int Splitter::GetOffset() const
{
    RECT bounds = { 0, 0, 0, 0 };

    GetBounds( &bounds );

    return bounds.left;
}

//===============================================================================================//
//  Description:
//      Get the offset of the divider window measured from the
//      right (or bottom) of the window
//
//  Parameters:
//      None
//
//  Returns:
//      int of offset from right or bottom
//===============================================================================================//
int Splitter::GetOffsetTwo() const
{
    int offsetTwo = 0;

    if ( m_hWindow && m_hWndOne && m_hWndTwo )
    {
        RECT parent = { 0, 0, 0, 0 };
        int  offset = GetOffset();

        // Get the size of the parent
        if ( GetClientRect( GetParent( m_hWindow ), &parent ) )
        {
            // Make sure it look sensible
            if ( parent.right > ( offset + DIVIDER_THICKNESS ) )
            {
                offsetTwo = parent.right - ( offset + DIVIDER_THICKNESS );
            }
        }
    }

    return offsetTwo;
}

//===============================================================================================//
//  Description:
//      Set the offset of the window from left or top depending
//      on the orientation
//
//  Parameters:
//      offset - the splitter's offset
//
//  Remarks:
//
//
//  Returns:
//      void
//===============================================================================================//
void Splitter::SetOffset( int offset )
{
    // If have create this window and have attached panels
    // to it them move them to the new offset
    if ( m_hWindow && m_hWndOne && m_hWndTwo )
    {
        RECT parent     = { 0, 0, 0, 0 };
        HWND hWndParent = GetParent( m_hWindow );

        if ( GetClientRect( hWndParent, &parent ) )
        {
            LONG W = parent.right  - parent.left;
            LONG H = parent.bottom - parent.top;

            MoveWindow( m_hWndOne, 0, 0, offset, H, TRUE );
            MoveWindow( m_hWindow, offset, 0, DIVIDER_THICKNESS, H, TRUE );
            MoveWindow( m_hWndTwo,
                        offset + DIVIDER_THICKNESS,
                        0, W - offset - DIVIDER_THICKNESS, H, TRUE );
        }
    }
}

//===============================================================================================//
//  Description:
//      Set the offset of the window from the right or bottom depending
//      on the orientation
//
//  Parameters:
//      offsetTwo - the offset
//
//  Remarks:
//
//      ------------------------
//      |         |S|           |
//      |         |p|           |
//      |         |l|           |
//      |   ONE   |i|    TWO    |
//      |         |t|           |
//      |         |t|           |
//      |         |e|           |
//      |         |r|           |
//      ------------------------
//      < offset >  < offsetTwo >
//
//  Returns:
//      void
//===============================================================================================//
void Splitter::SetOffsetTwo( int offsetTwo )
{
    int   offset = 0;

    // Calculate the Offset from the left
    if ( m_hWindow && m_hWndOne && m_hWndTwo )
    {
        // Get the size of the parent
        RECT  parent     = { 0, 0, 0, 0 };
        HWND  hWndParent = GetParent( m_hWindow );

        if ( GetClientRect( hWndParent, &parent ) )
        {
            if ( parent.right > ( offsetTwo + DIVIDER_THICKNESS ) )
            {
                offset = parent.right - ( offsetTwo + DIVIDER_THICKNESS );
            }
        }
        SetOffset( offset );
    }
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
//      void.
//
//===============================================================================================//
void Splitter::MouseLButtonDownEvent( const POINT& point, WPARAM /* keys */ )
{
    RECT    windowRect  = { 0, 0, 0, 0 };
    POINT   screenPoint = point;
    HCURSOR hCursor     = nullptr;

    if ( m_bPanelsAttached == false )
    {
        return;     // Nothing to do
    }

    hCursor = LoadCursor( nullptr, IDC_SIZEWE );
    if ( hCursor )
    {
        SetCursor( hCursor );
    }

    // Check if mouse was pressed inside, want screen coordinates of the cursor
    if ( GetWindowRect( m_hWindow, &windowRect   ) &&
         ClientToScreen( m_hWindow, &screenPoint ) &&
         PtInRect( &windowRect, screenPoint )       )
    {
        m_bMouseWasPressed = true;

        // Ensure splitter is on top
        SetWindowPos( m_hWindow,
                      nullptr,  // = HWND_TOP,
                      0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_SHOWWINDOW );
        SetCapture( m_hWindow );

        // Store points
        GetLocation( &m_SplitterLocation );
        m_InitialCursorLocation = screenPoint;
    }
}

//===============================================================================================//
//  Description:
//      Handle WM_LBUTTONUP event.
//
//  Parameters:
//      point - a POINT specifying the cursor position in the client area
//
//  Returns:
//      void.
//===============================================================================================//
void Splitter::MouseLButtonUpEvent( const POINT& /* point */ )
{
    WPARAM wParam = 0;

    if ( m_bPanelsAttached == false )
    {
        return;     // Nothing to do
    }
    m_bMouseWasPressed = false;
    ReleaseCapture();
    DoLayout();

    // Inform the listener in case its needs to know the splitter have moved
    if ( m_hWndAppMessageListener )
    {
        // For WM_COMMAND the control identifier is the low word of WPARAM
        wParam = MAKEWPARAM( PXS_APP_MSG_SPLITTER, 0 );
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
//      void.
//===============================================================================================//
void Splitter::MouseMoveEvent( const POINT& point, WPARAM /* keys */ )
{
    LONG    x = 0;
    RECT    parentBounds= { 0, 0, 0, 0 };
    POINT   screenPoint = point;
    HWND    hParent     = nullptr;
    HCURSOR hCursor     = nullptr;

    // Ensure panels attached and window handles are valid
    if ( m_bPanelsAttached == false ) return;

    hCursor = LoadCursor( nullptr, IDC_SIZEWE );
    if ( hCursor )
    {
        SetCursor( hCursor );
    }

    // If mouse was pressed, track its movements, want screen co-ordinates
    if ( m_bMouseWasPressed )
    {
        hParent = GetParent( m_hWindow );
        if ( hParent &&
             GetWindowRect( hParent, &parentBounds ) &&
             ClientToScreen( m_hWindow, &screenPoint )  )
        {
            x = screenPoint.x
                - m_InitialCursorLocation.x + m_SplitterLocation.x;
            x = PXSMaxInt( x, 20 );
            x = PXSMinInt( x, parentBounds.right - parentBounds.left
                                                 - (30 - DIVIDER_THICKNESS) );
            if ( x < 0 )
            {
                x = 0;
            }
            MoveWindow( m_hWindow,
                        x, 0, DIVIDER_THICKNESS, parentBounds.bottom - parentBounds.top, TRUE );
        }
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
void Splitter::PaintEvent( HDC hdc )
{
    int     i = 0, nStart  = 0, nEnd = 0;
    RECT    recDot = { 0, 0, 0, 0 }, recClient = { 0, 0, 0, 0 };
    HPEN    hPen  = nullptr;
    POINT   ptTriangle[ 3 ];
    HBRUSH  hBrush       = nullptr, hbrWhite   = nullptr;
    HGDIOBJ hObjBrushOld = nullptr, hObjPenOld = nullptr;

    if ( ( hdc == nullptr ) || ( m_hWindow == nullptr ) )
    {
        return;     // Nothing to do
    }
    DrawBackground( hdc );
    GetClientRect( m_hWindow, &recClient );

    // Create pen and brushes
    hbrWhite = (HBRUSH)GetStockObject( WHITE_BRUSH );
    hPen     = CreatePen( PS_SOLID, 1, m_crForeground );
    hBrush   = CreateSolidBrush( m_crForeground );
    if ( hPen && hBrush )
    {
        hObjPenOld   = SelectObject( hdc, hPen );
        hObjBrushOld = SelectObject( hdc, hBrush );

        nStart = ( 2 * (recClient.bottom - recClient.top ) ) / 5;
        nEnd   = ( 3 * (recClient.bottom - recClient.top ) ) / 5;

        // Draw first triangle
        ptTriangle [ 0 ].x = recClient.left + 2;
        ptTriangle [ 0 ].y = nStart - 15;
        ptTriangle [ 1 ].x = ptTriangle [ 0 ].x + 4;
        ptTriangle [ 1 ].y = ptTriangle [ 0 ].y + 4;
        ptTriangle [ 2 ].x = ptTriangle [ 0 ].x;
        ptTriangle [ 2 ].y = ptTriangle [ 0 ].y + 8;
        Polygon( hdc, ptTriangle, 3 );

        // Draw second triangle
        ptTriangle [ 0 ].x = recClient.left + 2;
        ptTriangle [ 0 ].y = nEnd + 8;
        ptTriangle [ 1 ].x = ptTriangle [ 0 ].x + 4;
        ptTriangle [ 1 ].y = ptTriangle [ 0 ].y - 4;
        ptTriangle [ 2 ].x = ptTriangle [ 0 ].x + 4;
        ptTriangle [ 2 ].y = ptTriangle [ 0 ].y + 4;
        Polygon( hdc, ptTriangle, ARRAYSIZE( ptTriangle ) );

        // Draw 2 columns of dots in middle 1/5
        for ( i = nStart; i < nEnd; i += 3 )
        {
            // Set dot's position
            recDot.left   = recClient.left + 2;
            recDot.top    = i;
            recDot.right  = recDot.left + 1;
            recDot.bottom = recDot.top  + 1;
            FillRect( hdc, &recDot, hBrush );

            // Fill shadow
            OffsetRect( &recDot, 1, 1 );
            FillRect( hdc, &recDot, hbrWhite );

            // Second column
            OffsetRect( &recDot, 2, -1 );
            FillRect( hdc, &recDot, hBrush );

            // Fill shadow
            OffsetRect( &recDot, 1, 1 );
            FillRect( hdc, &recDot, hbrWhite );
        }

        // Reset the device context
        if ( hObjPenOld   ) SelectObject( hdc, hObjPenOld );
        if ( hObjBrushOld ) SelectObject( hdc, hObjBrushOld );
    }
    if ( hPen   ) DeleteObject( hPen );
    if ( hBrush ) DeleteObject( hBrush );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
