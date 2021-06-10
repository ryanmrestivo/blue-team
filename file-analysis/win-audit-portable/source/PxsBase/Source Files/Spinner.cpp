///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Spinner Control Class Implementation
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
#include "PxsBase/Header Files/Spinner.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StaticControl.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
Spinner::Spinner()
        :TIMER_ID( 1010 ),
         TIMER_INTERVAL ( 100 ),
         WIDTH_BUTTONS( 25 ),
         m_bMouseWasPressed( false ),
         m_bFirstTimerEvent( false ),
         m_uHigh( 100 ),
         m_uLow( 0 ),
         m_uValue( 0 ),
         m_TextField()
{
    // Set properties. Do not set WS_EX_CONTROLPARENT style for this window
    // otherwise will get unwanted WM_GETDLGCODE events for the text field if
    // there are also radio buttons on the dialog.
    m_CreateStruct.style  = WS_CHILD | WS_VISIBLE | WS_CLIPCHILDREN;
    try
    {
        SetStyle( WS_TABSTOP, true );
        SetDoubleBuffered( true );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy constructor - not allowed so no implementation

// Destructor
Spinner::~Spinner()
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
//      Create this window
//
//  Parameters:
//      hWndParent - handle to the the parent window
//
//  Returns:
//      Zero
//===============================================================================================//
int Spinner::Create( HWND hWndParent )
{
    SIZE      windowSize = { 0, 0 };
    String    Text;
    Formatter Format;

    // Call the super's method
    Window::Create( hWndParent );

    // Create the edit box, Only allow creation once
    // do not set the initial value in the box in case
    // NULL is a valid value, i.e. a default
    if ( m_hWindow )
    {
        GetSize( &windowSize );
        m_TextField.SetDigitsOnly();
        m_TextField.SetBounds( 0, 0, windowSize.cx - WIDTH_BUTTONS, windowSize.cy );
        m_TextField.Create( m_hWindow );

        // Set the text and the maximum characters
        Text = Format.UInt32( m_uValue );
        m_TextField.SetText( Text );

        Text = Format.UInt32( m_uHigh );
        m_TextField.SetTextMaxLength( Text.GetLength() );
    }

    return 0;
}

//===============================================================================================//
//  Description:
//      Get the range for this up/down control
//
//  Parameters:
//      pLow  - receives the lower bound
//      pHigh - receives the upper bound
//
//  Returns:
//      void
//===============================================================================================//
void Spinner::GetRange( DWORD* pLow, DWORD* pHigh )
{
    if ( ( pLow == nullptr ) || ( pHigh == nullptr ) )
    {
        throw ParameterException( L"pLow/pHigh", __FUNCTION__ );
    }
    *pLow  = m_uLow;
    *pHigh = m_uHigh;
}

//===============================================================================================//
//  Description:
//      Get the value of this up/down control
//
//  Parameters:
//      None
//
//  Returns:
//      DWORD
//===============================================================================================//
DWORD Spinner::GetValue()
{
    return m_uValue;
}

//===============================================================================================//
//  Description:
//      Sets the window to en/disabled to mouse and keyboard input
//
//  Parameters:
//        enabled - flag to set window enabled or disabled
//
//  Returns:
//      void
//===============================================================================================//
void Spinner::SetEnabled( bool enabled )
{
    if ( enabled )
    {
        EnableWindow( m_hWindow, TRUE );
    }
    else
    {
        EnableWindow( m_hWindow, FALSE );
    }
    m_TextField.SetEnabled( enabled );
    Repaint();
}

//===============================================================================================//
//  Description:
//      Set the range for this up/down control
//
//  Parameters:
//      low  - the lower bound
//      high - the upper bound
//
//  Returns:
//      void
//===============================================================================================//
void Spinner::SetRange( DWORD low, DWORD high )
{
    size_t    maxLen = 0;
    String    LowString, HighString;
    Formatter Format;

    // Make sure inputs in order
    if ( high >= low )
    {
        m_uLow  = low;
        m_uHigh = high;
    }
    else
    {
        // Reverse them
        m_uLow  = high;
        m_uHigh = low;
    }

    // The existing value may now be out of bounds so 'set' it
    // to ensure its in range
    SetValue( m_uValue );

    // Use the high range to set the maximum number of characters
    LowString  = Format.UInt32( m_uLow  );
    HighString = Format.UInt32( m_uHigh );
    maxLen = PXSMaxSizeT( LowString.GetLength(), HighString.GetLength() );
    m_TextField.SetTextMaxLength( maxLen );
}

//===============================================================================================//
//  Description:
//      Set the value of this up/down control
//
//  Parameters:
//      value - the new value
//
//  Remarks:
//      Keep the text box in sync, note this raises an EN_CHANGE event
//      so ensure it does not cause an infinite cycle
//
//  Returns:
//      void
//===============================================================================================//
void Spinner::SetValue( DWORD value )
{
    Formatter Format;

    // See if there is anything to do.
    if ( m_uValue == value ) return;

    // Make sure in range
    if ( value > m_uHigh )
    {
        m_uValue = m_uHigh;
    }
    else if ( value < m_uLow )
    {
        m_uValue = m_uLow;
    }
    else
    {
        // OK, in range
        m_uValue = value;
    }
    m_TextField.SetText( Format.UInt32( m_uValue ) );
    SendMessage( GetParent( m_hWindow ),
                            WM_COMMAND, MAKEWPARAM( 0, BN_CLICKED ), (LPARAM)m_hWindow );
}

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
//      0 if handled, else non-zero
//===============================================================================================//
LRESULT Spinner::CommandEvent( WPARAM wParam, LPARAM /* lParam */ )
{
    DWORD    value  = 0;
    wchar_t* endptr = nullptr;
    LPCWSTR  psz    = nullptr;
    String   Text;

    if ( HIWORD( wParam ) == EN_CHANGE )
    {
        m_TextField.GetText( &Text );
        psz = Text.c_str();
        if ( psz )
        {
            value = wcstoul( psz, &endptr, 10 );
            if ( value != m_uValue )
            {
                // Only process if in bounds
                if ( ( value >= m_uLow ) && ( value <= m_uHigh ) )
                {
                    m_uValue = value;

                    // Inform the parent via a button click
                    SendMessage( GetParent( m_hWindow ),
                                 WM_COMMAND, MAKEWPARAM( 0, BN_CLICKED ), (LPARAM)m_hWindow );
                }
            }
        }
    }
    else if ( HIWORD( wParam ) == EN_KILLFOCUS )
    {
        m_TextField.GetText( &Text );
        psz = Text.c_str();
        if ( psz )
        {
            value = wcstoul( psz, &endptr, 10 );
            if ( value != m_uValue )
            {
                // Limit
                PXSLimitUInt32( m_uLow, m_uHigh, &value );
                SetValue( value );
            }
        }
    }

    return 0;   // Handled
}

//===============================================================================================//
//  Description:
//      Handle WM_SETFOCUS event for this window
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void Spinner::GainFocusEvent()
{
    HWND hWndTextField = m_TextField.GetHwnd();

    // Forward the focus to the text box and select all its contents
    if ( hWndTextField )
    {
        if ( hWndTextField != GetFocus() )
        {
            SetFocus( hWndTextField );
        }
        SendMessage( hWndTextField, EM_SETSEL, 0, -1 );
    }
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
void Spinner::MouseLButtonDblClickEvent( const POINT& point )
{
    HandleValueChangeEvent( point );
}

//===============================================================================================//
//  Description:
//      Handle WM_LBUTTONDOWN event for this window
//
//  Parameters:
//      point - POINT in windows where mouse was clicked in the client area
//      keys  - which virtual keys are down
//
//  Returns:
//      void
//===============================================================================================//
void Spinner::MouseLButtonDownEvent( const POINT& point, WPARAM /* keys */ )
{
    if ( m_hWindow == nullptr )
    {
        return;     // Nothing to do
    }

    m_bMouseWasPressed = true;
    if ( m_hWindow != GetFocus() )
    {
        SetFocus( m_hWindow );
    }
    SetCapture( m_hWindow );
    HandleValueChangeEvent( point );

    // Start the timer with a 1/2 second interval to auto change increment
    // the value if the mouse is held down for over 1/2 second
    m_bFirstTimerEvent = true;
    SetTimer( m_hWindow, TIMER_ID, 500, nullptr );
}

//===============================================================================================//
//  Description:
//      Handle WM_LBUTTONUP event for this window
//
//  Parameters:
//      point - POINT in windows where mouse was clicked in the client area
//
//  Returns:
//      void
//===============================================================================================//
void Spinner::MouseLButtonUpEvent( const POINT& /* point */ )
{
    m_bMouseWasPressed = false;
    ReleaseCapture();
    KillTimer( m_hWindow, TIMER_ID );
    Repaint();
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
void Spinner::PaintEvent( HDC hdc )
{
    RECT     halfRect   = { 0, 0, 0, 0 };
    RECT     clientRect = { 0, 0, 0, 0 };
    POINT    cursorPos  = { 0, 0 }, centrePoint = { 0, 0 };
    DWORD    i = 0;
    HGDIOBJ  originalPen = nullptr;
    COLORREF penColour   = m_crForeground;
    StaticControl Static;

    if ( hdc == nullptr )
    {
        return;
    }
    DrawBackground( hdc );
    GetClientRect( m_hWindow, &clientRect );
    GetCursorPos( &cursorPos );
    ScreenToClient( m_hWindow, &cursorPos );

    originalPen = SelectObject( hdc, GetStockObject( DC_PEN ) );
    if ( IsEnabled() == false )
    {
        penColour = GetSysColor( COLOR_GRAYTEXT );
    }
    SetDCPenColor( hdc, penColour );

    // Set upper and lower rectangles
    for ( i = 0; i < 2; i++ )
    {
        halfRect.left  = clientRect.right - WIDTH_BUTTONS;
        halfRect.right = clientRect.right;
        if ( i == 0 )
        {
            halfRect.top    = 1;
            halfRect.bottom = ( clientRect.bottom / 2 );
        }
        else
        {
            halfRect.top    = ( clientRect.bottom / 2 );
            halfRect.bottom = clientRect.bottom - 1;
        }

        Static.SetBounds( halfRect );
        Static.SetBackground( RGB( 221, 221, 221 ) );
        Static.SetShape( PXS_SHAPE_RAISED );
        if ( PtInRect( &halfRect, cursorPos ) )
        {
            if ( m_bMouseWasPressed )
            {
                Static.SetShape( PXS_SHAPE_SUNK );
            }
        }
        Static.Draw( hdc );
        Static.Reset();

        // Draw a + or a -, work out the centre of the button
        centrePoint.x = halfRect.left + ( (halfRect.right  - halfRect.left)/2 );
        centrePoint.y = halfRect.top  + ( (halfRect.bottom - halfRect.top )/2 );

        MoveToEx( hdc, centrePoint.x - 2, centrePoint.y, nullptr );
        LineTo(   hdc, centrePoint.x + 2, centrePoint.y );
        if ( i == 0 )
        {
            // Vertical line of plus
            MoveToEx( hdc, centrePoint.x, centrePoint.y - 2, nullptr );
            LineTo(   hdc, centrePoint.x, centrePoint.y + 3 );
        }
    }
    m_TextField.Repaint();

    // Reset DC
    if ( originalPen )
    {
        SelectObject( hdc, originalPen );
    }
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
void Spinner::TimerEvent( UINT_PTR /* timerID */ )
{
    POINT cursorPos = { 0, 0 };

    // If this is the first timer event, reduce the interval so that the next
    // event arrives at the expected interval. This is because on the first
    // mouse click there is a short delay before the value starts changing
    // automatically.
    if ( m_bFirstTimerEvent )
    {
        m_bFirstTimerEvent = false;
        KillTimer( m_hWindow, TIMER_ID );
        SetTimer( m_hWindow, TIMER_ID, TIMER_INTERVAL, nullptr );
    }

    GetCursorPos( &cursorPos );
    HandleValueChangeEvent( cursorPos );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Handle the effect of a change of value on this control.
//
//  Parameters:
//      point - position of the cursor in the client area
//
//  Returns:
//      void
//===============================================================================================//
void Spinner::HandleValueChangeEvent( const POINT& point )
{
    DWORD newValue   = 0;
    RECT  clientRect = { 0, 0, 0, 0 };
    RECT  buttonRect = { 0, 0, 0, 0 };
    Formatter Format;

    if ( m_hWindow == nullptr )
    {
        return;     // Nothing to do
    }
    GetClientRect( m_hWindow, &clientRect );

    // Only change the value if the mouse is in the button area
    buttonRect = clientRect;
    buttonRect.left = buttonRect.right - WIDTH_BUTTONS;
    if ( ( PtInRect( &buttonRect, point )     ) &&
         ( buttonRect.right > buttonRect.left )  )
    {
        newValue = m_uValue;

        // If mouse in top half then increment, else decrement. Account for
        // border thickness so that the buttons as drawn in a pushed state
        if ( ( point.y >= 1 ) && ( point.y < ( clientRect.bottom / 2 ) ) )
        {
            newValue++;
        }
        else if ( point.y > ( clientRect.bottom / 2 ) &&
                ( point.y < ( clientRect.bottom - 1 ) ) )
        {
            newValue--;
        }
        SetValue( newValue );
    }
    Repaint();
}
