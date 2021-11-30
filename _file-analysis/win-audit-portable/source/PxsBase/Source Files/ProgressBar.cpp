///////////////////////////////////////////////////////////////////////////////////////////////////
//
//  Progress Bar Class Implementation
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
#include "PxsBase/Header Files/ProgressBar.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/FunctionException.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StaticControl.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
ProgressBar::ProgressBar()
            :TIMER_ID( 1003 ) ,
             m_uMaximum( 100 ),
             m_uStep( 5 ),
             m_uValue( 0 ),
             m_Note()

{
    // Creation parameters, initially invisible
    m_CreateStruct.cy    = 20;
    m_CreateStruct.cx    = 200;
    m_CreateStruct.style = WS_CHILD;

    // Properties
    try
    {
        SetBorderStyle( PXS_SHAPE_RECTANGLE_DOTTED );
        SetDoubleBuffered( true );
        SetBackground( GetSysColor( COLOR_BTNFACE ) );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}


// Copy constructor - no allowed so no implementation

// Destructor
ProgressBar::~ProgressBar()
{
    // Clean up
    if ( m_hWindow )
    {
        KillTimer( m_hWindow, TIMER_ID );  // Non-op if no timer
    }
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Operators
///////////////////////////////////////////////////////////////////////////////////////////////////

// Assignment operator - no allowed so no implementation

///////////////////////////////////////////////////////////////////////////////////////////////////
// Public Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Redraw the progress bar now.
//
//  Parameters:
//      None
//
//  Remarks:
//      Redraws the progress bar immediately reflecting any change in its
//      value.
//
//  Returns:
//      void
//===============================================================================================//
void ProgressBar::RedrawNow()
{
    if ( !m_hWindow )
    {
        return;     // Nothing to do
    }

    // Process the message queue so that the timer event is handled,
    // this ensures the value of the progress bar is updated
    DoPaintEvents( TIMER_ID );
    RedrawWindow( m_hWindow,
                  nullptr, nullptr, RDW_ERASE | RDW_INVALIDATE | RDW_UPDATENOW | RDW_ALLCHILDREN);
}

//===============================================================================================//
//  Description:
//      Set the text shown as the bar progresses. This is shown centre aligned
//      in the coloured part of the progress bar.
//
//  Parameters:
//      Note - the text
//
//  Returns:
//      void
//===============================================================================================//
void ProgressBar::SetNote( const String& Note )
{
    m_Note = Note;
    RedrawNow();
}

//===============================================================================================//
//  Description:
//      Set the percentage value of the progress bar
//
//  Parameters:
//      percentage - value of the progress bar
//
//  Returns:
//      void
//===============================================================================================//
void ProgressBar::SetPercentage( DWORD percentage )
{
    if ( percentage > 100 )
    {
        percentage = 100;
    }

    m_uValue = ( percentage * m_uMaximum ) / 100;
    RedrawNow();
}


//===============================================================================================//
//  Description:
//      Start the cyclic progression
//
//  Parameters:
//      timerInterval - update time interval of progress bar in milli-seconds
//
//  Returns:
//      void
//===============================================================================================//
void ProgressBar::Start( UINT timerInterval )
{
    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    if ( timerInterval == 0 )
    {
        throw ParameterException( L"timerInterval", __FUNCTION__ );
    }

    m_uValue = 0;
    SetTimer( m_hWindow, TIMER_ID, timerInterval, nullptr );
    SetVisible( true );
    Repaint();
}

//===============================================================================================//
//  Description:
//      Stop the cyclic progression
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void ProgressBar::Stop()
{
    if ( m_hWindow )
    {
        KillTimer( m_hWindow, TIMER_ID );
    }
    m_uValue = 0;
    SetVisible( false );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

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
void ProgressBar::PaintEvent( HDC hdc )
{
    int  W = 0, value = 0;
    RECT bounds = { 0, 0, 0, 0 };
    StaticControl Static;

    if ( ( hdc == nullptr ) || ( m_hWindow == nullptr ) )
    {
        return;     // Nothing to do
    }
    DrawBackground( hdc );

    GetClientRect( m_hWindow, &bounds );
    if ( m_uValue < INT_MAX )
    {
        value = PXSCastUInt32ToInt32( m_uValue );
    }

    W = bounds.right - bounds.left;
    if ( W > 0 )
    {
        // Note using the foreground colour to draw the progress bar's value
        Static.SetBackground( m_crForeground );
        if ( IsRightToLeftReading() )
        {
            Static.SetAlignmentX( PXS_RIGHT_ALIGNMENT );
            bounds.left = bounds.right - ( ( value * W ) / 100 );
            Static.SetBackgroundGradient( PXS_COLOUR_WHITE, m_crForeground, false );
        }
        else
        {
            Static.SetAlignmentX( PXS_LEFT_ALIGNMENT );
            bounds.right = (value * W ) / 100;
            Static.SetBackgroundGradient( m_crForeground, PXS_COLOUR_WHITE, false );
        }
        Static.SetBounds( bounds );
        Static.SetShape( m_uBorderStyle );
        Static.SetShapeColour( m_crForeground );
        Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
        Static.SetSingleLineText( m_Note );
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
//  Returns:
//      void
//===============================================================================================//
void ProgressBar::TimerEvent( UINT_PTR /* timerID */ )
{
    // Start over if got to the maximum
    if ( m_uValue > ( m_uMaximum - m_uStep ) )
    {
        m_uValue = 0;
    }
    else
    {
        m_uValue = m_uValue + m_uStep;
    }
    Repaint();
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
