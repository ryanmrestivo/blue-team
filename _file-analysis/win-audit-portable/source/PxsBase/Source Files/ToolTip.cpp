///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Tool Tip Class Implementation
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
#include "PxsBase/Header Files/ToolTip.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/NullException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Static declaration of the one and only instance of this object
ToolTip* ToolTip::m_pInstance = nullptr;

// Default constructor - private as singleton class

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator - not allowed so no implementation

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Destroy an instance of this class.
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void ToolTip::DestroyInstance()
{
    if ( m_pInstance )
    {
        // Destroy the window first then this instance
        if ( m_pInstance->GetHwnd() )
        {
            DestroyWindow( m_pInstance->GetHwnd() );
        }
        delete m_pInstance;
        m_pInstance = nullptr;
    }
}

//===============================================================================================//
//  Description:
//      Get an instance of this class.
//
//  Parameters:
//      None
//
//  Returns:
//      Pointer to this, NULL if cannot create an instance
//===============================================================================================//
ToolTip* ToolTip::GetInstance()
{
    HWND hWndMainFrame;

    if ( m_pInstance )
    {
        return m_pInstance;
    }

    m_pInstance = new ToolTip;
    if ( m_pInstance == nullptr )
    {
        throw MemoryException( __FUNCTION__ );
    }

    // Must have allocated the application object
    if ( g_pApplication == nullptr )
    {
        return nullptr;
    }
    hWndMainFrame = g_pApplication->GetHwndMainFrame();
    if ( hWndMainFrame )
    {
        m_pInstance->Create( hWndMainFrame );
        SetParent( m_pInstance->GetHwnd(), GetDesktopWindow() );
    }

    return m_pInstance;
}

//===============================================================================================//
//  Description:
//      Get this tool tip's text
//
//  Parameters:
//      none
//
//  Remarks:
//      This class overrides the Window::SetText method so need to ensure
//      GetText retrieves that value otherwise will return the empty
//      string stored in the window.
//
//  Returns:
//     Reference to the text
//===============================================================================================//
const String& ToolTip::GetText() const
{
    return m_Tip.GetText();
}

//===============================================================================================//
//  Description:
//      Determine if an instance of this singleton has been created
//
//  Parameters:
//      none
//
//  Returns:
//     true if intance created, otherwise false
//===============================================================================================//
bool ToolTip::IsInstanceCreated()
{
    if ( m_pInstance )
    {
        return true;
    }

    return false;
}

//===============================================================================================//
//  Description:
//      Tell the tool tip that a mouse down event occurred
//
//  Parameters:
//      point - position of cursor in the owner's client rectangle
//
//  Remarks:
//      Will not show the tip once the mouse has been pressed
//      inside the owner rectangle
//
//  Returns:
//     void
//===============================================================================================//
void ToolTip::MouseButtonDown( const POINT& point )
{
    // Verify the mouse button was pressed in the owner's rectangle
    if ( PtInRect( &m_OwnerBounds, point ) )
    {
        m_bMousePressed = true;
        HideTip();
    }
}

//===============================================================================================//
//  Description:
//      Set the display time in milliseconds
//
//  Parameters:
//      dismissDelay - time to display the tool tip in milliseconds
//
//  Returns:
//      void
//===============================================================================================//
void ToolTip::SetDismissDelay( DWORD dismissDelay )
{
    m_uDismissDelay = dismissDelay;
}

//===============================================================================================//
//  Description:
//      Set the delay in milliseconds before the tool tip is shown
//
//  Parameters:
//      initialDelay - delay in milliseconds
//
//  Returns:
//      void
//===============================================================================================//
void ToolTip::SetInitialDelay( DWORD initialDelay )
{
    m_uInitialDelay = initialDelay;
}

//===============================================================================================//
//  Description:
//      Hide the tool tip
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void ToolTip::HideTip()
{
    if ( m_hWindow )
    {
        KillTimer( m_hWindow, TIMER_ID );
    }
    SetVisible( false );
    m_bMouseInside  = false;
    m_bMousePressed = false;
}

//===============================================================================================//
//  Description:
//      Show the tool tip
//
//  Parameters:
//      TipText - the text to display
//      owner   - tool tip's owner rectangle in screen coordinates
//
//  Returns:
//      void
//===============================================================================================//
void ToolTip::Show( const String& TipText, const RECT& owner )
{
    HDC hdc = nullptr;

    if ( m_hWindow == nullptr )
    {
        return;     // No reason to show the tool tip
    }

    // If the input rectangle is not the same as the previous
    // one then reset the mouse flags
    if ( EqualRect( &owner, &m_OwnerBounds ) == 0 )
    {
        m_bMouseInside  = false;
        m_bMousePressed = false;
    }

    // Test if mouse is already inside
    if ( m_bMouseInside )
    {
        return;     // Nothing to do.
    }
    m_bMouseInside = true;
    HideTip();

    // Get the device context of the desktop
    hdc = GetDC( nullptr );
    if ( hdc )
    {
        // Don't allow this method to throw
        try
        {
            m_OwnerBounds = owner;
            SetText( TipText );

            // Start the timer. If the delay time is short, i.e. less than
            // the timer interval, call the timer method now so don't have
            // to wait for the first callback
            if ( m_uInitialDelay <= TIMER_INTERVAL )
            {
                TimerEvent( TIMER_ID );
            }
            SetTimer( m_hWindow, TIMER_ID, TIMER_INTERVAL, nullptr );
            m_uElapsedTime = 0;
        }
        catch ( const Exception& e )
        {
            PXSLogException( e, __FUNCTION__ );
        }
        ReleaseDC( nullptr, hdc );
    }
}

//===============================================================================================//
//  Description:
//      Set the text of the tool tip
//
//  Parameters:
//      TipText - the text to display
//
//  Remarks:
//     Recalculate the size of the window required to hold the text
//
//      Tool tips are optional, do not let exceptions propagate out of
//      this method
//
//  Returns:
//      void
//===============================================================================================//
void ToolTip::SetText( const String& Text )
{
    RECT    bounds   = { 0, 0, 0, 0 };
    SIZE    textSize = { 0, 0 };
    HDC     hdc      = nullptr;
    HFONT   hFont    = nullptr;
    HGDIOBJ oldFont  = nullptr;
    LOGFONT logFont;

    m_Tip.SetText( Text );
    if ( Text.IsEmpty() )
    {
        return;     // No tip to draw
    }

    // Class scope checks
    if ( m_hWindow == nullptr )
    {
        return;     // Nothing to do
    }

    hdc = GetDC( m_hWindow );
    if ( hdc == nullptr )
    {
        return;     // Nothing to draw on
    }

    // Need to release the DC
    try
    {
        memset( &logFont, 0, sizeof ( logFont ) );
        m_Tip.GetFont().GetLogFont( &logFont );
        hFont = CreateFontIndirect( &logFont );
        if ( hFont )
        {
            oldFont = SelectObject( hdc, hFont );
            GetTextExtentPoint32( hdc, Text.c_str(), lstrlen( Text.c_str() ), &textSize );
            // Reset
            if ( oldFont )
            {
                SelectObject( hdc, oldFont );
            }
            DeleteObject( hFont );
        }

        //  Set the window size, add padding
        textSize.cx += ( 3 * m_Tip.GetPadding() );
        textSize.cy += ( 2 * m_Tip.GetPadding() );
        SetSize( textSize );

        // Set bounds for static border shape
        bounds.left    = 0;
        bounds.top     = 0;
        bounds.right   = bounds.left + textSize.cx;
        bounds.bottom  = bounds.top  + textSize.cy;
        m_Tip.SetBounds( bounds );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
    ReleaseDC( nullptr, hdc );
    Repaint();
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Paint event handler
//
//  Parameters:
//        hdc - Handle to device context
//
//  Remarks:
//      Tool tips are optional, do not let exceptions propagate out of
//      this method
//
//  Returns:
//      void.
//===============================================================================================//
void ToolTip::PaintEvent( HDC hdc )
{
    if ( ( hdc == nullptr ) || ( m_hWindow == nullptr ) )
    {
        return;
    }

    m_Tip.SetBackground( m_crBackground );
    m_Tip.SetForeground( m_crForeground );
    m_Tip.SetBackgroundGradient( m_crGradient1, m_crGradient2, false );
    m_Tip.Draw( hdc );
}

//===============================================================================================//
//  Description:
//      Handle WM_TIMER event.
//
//  Parameters:
//      timerID - The identifier of the timer that fired this event
//
//  Remarks:
//      Tool tips are optional so no exceptions are raised.
//
//  Returns:
//      void
//===============================================================================================//
void ToolTip::TimerEvent( UINT_PTR /* timerID */ )
{
    int   xPos = 0, yPos = 0;
    RECT  clientRect = { 0, 0, 0, 0 }, workArea = { 0, 0, 0, 0 };
    POINT cursorPos  = { 0, 0 };

    m_uElapsedTime += TIMER_INTERVAL;
    if ( m_hWindow == nullptr )
    {
        return;     // No window, should get here as the window owns the timer
    }
    GetCursorPos( &cursorPos );
    GetClientRect( m_hWindow, &clientRect );
    SystemParametersInfo( SPI_GETWORKAREA, 0, &workArea, 0 );

    try
    {
        // Is the cursor still in the owner's rectangle
        if ( PtInRect( &m_OwnerBounds, cursorPos ) )
        {
            m_bMouseInside = true;

            // Decide what to do based on elapsed time
            if ( ( m_uElapsedTime >=  m_uInitialDelay ) &&
                 ( m_uElapsedTime < ( m_uInitialDelay + m_uDismissDelay ) ) )
            {
                // Show the tool tip, if the mouse was not pressed and there
                // is some text to display
                if ( ( m_bMousePressed == false ) &&
                     ( m_Tip.GetText().GetLength() ) )
                {
                    // Limit the position to the work area dimensions.
                    xPos = cursorPos.x;
                    yPos = cursorPos.y + 20;
                    xPos = PXSMinInt( xPos, workArea.right - clientRect.right );
                    xPos = PXSMaxInt( xPos, 0 );
                    yPos = PXSMinInt( yPos, workArea.bottom-clientRect.bottom );
                    yPos = PXSMaxInt( yPos, 0 );
                    SetWindowPos( m_hWindow,
                                  nullptr,  // = HWND_TOP,
                                  xPos, yPos, 0, 0, SWP_NOSIZE | SWP_SHOWWINDOW );
                }
            }
            else if ( m_uElapsedTime >= m_uDismissDelay )
            {
                // Tool tip displayed for allotted time, so hide it
                KillTimer( m_hWindow, TIMER_ID );
                HideTip();
            }
        }
        else
        {
            // Cursor has moved out of rectangle
            m_bMouseInside  = false;
            m_bMousePressed = false;
            KillTimer( m_hWindow, TIMER_ID );
            HideTip();
        }
    }
    catch ( const Exception& )
    {
        // Something went wrong, stop the timer
        KillTimer( m_hWindow, TIMER_ID );
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor - private as singleton class
ToolTip::ToolTip()
        :m_bMouseInside( false ),
         m_bMousePressed( false ),
         m_uInitialDelay( 0 ),
         m_uDismissDelay( 0 ),
         m_uElapsedTime( 0 ),
         m_OwnerBounds( ),
         TIMER_ID ( 1002 ),
         TIMER_INTERVAL ( 100 ),
         m_Tip()
{
    // Creation parameters
    m_CreateStruct.style      = WS_CHILD | WS_CLIPSIBLINGS;     // Invisible
    m_CreateStruct.dwExStyle |= WS_EX_TOOLWINDOW;
    m_CreateStruct.cy         = 16;     // Nominal height
    m_CreateStruct.cx         = 50;     // Nominal width

    m_crBackground = GetSysColor( COLOR_INFOBK );
    m_crForeground = GetSysColor( COLOR_INFOTEXT );
    try
    {
        Init();
        SetDoubleBuffered( true );
        SetVisible( false );
        m_Tip.SetAlignmentY( PXS_CENTER_ALIGNMENT );
        m_Tip.SetPadding( 4 );
        m_Tip.SetShape( PXS_SHAPE_SHADOW );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy Constructor - not allowed so no implementation

// Destructor  - private as singleton class
ToolTip::~ToolTip()
{
    // Ensure the timer has been killed
    if ( m_hWindow )
    {
        KillTimer( m_hWindow, TIMER_ID );
    }
}

//===============================================================================================//
//  Description:
//      Reset the state of the tool tip
//
//  Parameters:
//      None
//
//  Remarks:
//      This is a Singleton class so any changes to its properties
//      can be reset in case another window want to use it.
//
//      Called by constructor so does not throw
//
//  Returns:
//      void
//===============================================================================================//
void ToolTip::Init()
{
    m_bMouseInside  = false;
    m_bMousePressed = false;
    m_uInitialDelay = 500;        // Half a second
    m_uDismissDelay = 4000;       // Four seconds
    m_uElapsedTime  = 0;
    memset( &m_OwnerBounds, 0, sizeof ( m_OwnerBounds ) );
    try
    {
        HideTip();
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}
