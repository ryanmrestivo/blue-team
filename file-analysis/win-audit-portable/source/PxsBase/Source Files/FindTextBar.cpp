///////////////////////////////////////////////////////////////////////////////
//
// Find Text Bar Implementation
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
///////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/FindTextBar.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/ParameterException.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
FindTextBar::FindTextBar()
            :FIND_TEXT_BAR_HEIGHT( 40 ),
             m_Find(),
             m_TextNotFound(),
             m_TextField(),
             m_CaseSensitive(),
             m_FindBackward( 23, 20, 16, 16 ),
             m_FindForward( 23, 20, 16, 16 )
{
    // Creation parameters, initialy invisible
    m_CreateStruct.cy    = FIND_TEXT_BAR_HEIGHT;
    m_CreateStruct.style = WS_BORDER | WS_CHILD | WS_CLIPCHILDREN | SS_OWNERDRAW;

    try
    {
        SetBackground( GetSysColor( COLOR_BTNFACE ) );
        SetLayoutProperties( PXS_LAYOUT_STYLE_ROW_MIDDLE, 5, 3 );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Destructor
FindTextBar::~FindTextBar()
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
//     Create the controls for this windows
//
//  Parameters:
//      hWndAppMessageListener - the window listening for application
//                               specific messages
//
//  Returns:
//      void
//===============================================================================================//
void FindTextBar::CreateControls( HWND hWndAppMessageListener )
{
    SIZE size = { 0, 0 };

    // One shot
    if ( m_bControlsCreated )
    {
        throw FunctionException( L"m_bControlsCreated", __FUNCTION__ );
    }

    if ( m_hWindow == nullptr )
    {
        throw FunctionException( L"m_hWindow", __FUNCTION__ );
    }

    // Word/Phrase label, will suppress WS_EX_TRANSPARENT as causes flicker
    // and/or paint problems when this window is resized
    m_Find.SetExStyle( WS_EX_TRANSPARENT, false );
    m_Find.SetBackground( m_crBackground );
    m_Find.Create( m_hWindow );
    m_Find.GetSize( &size );
    m_Find.SetSize( 120, size.cy );

    // Edit box
    m_TextField.Create( m_hWindow );
    m_TextField.GetSize( &size );
    m_TextField.SetSize( 200, size.cy );

    // Create the forward find image button
    m_FindForward.Create( m_hWindow );
    m_FindForward.SetDisplayCaption( false );
    m_FindForward.SetAppMessageCode( PXS_APP_MSG_BUTTON_CLICK );
    m_FindForward.SetAppMessageListener( hWndAppMessageListener );
    m_FindForward.SetBackground( m_crBackground );
    m_FindForward.SetFilledBitmap( PXS_COLOUR_BLUE, PXS_COLOUR_BLACK, PXS_SHAPE_ARROW_DOWN );

    // Backward find image button
    m_FindBackward.Create( m_hWindow );
    m_FindBackward.SetDisplayCaption( false );
    m_FindBackward.SetAppMessageCode( PXS_APP_MSG_BUTTON_CLICK );
    m_FindBackward.SetAppMessageListener( hWndAppMessageListener );
    m_FindBackward.SetBackground( m_crBackground );
    m_FindBackward.SetFilledBitmap( PXS_COLOUR_BLUE, PXS_COLOUR_BLACK, PXS_SHAPE_ARROW_UP );

    // Case sensitive check box
    m_CaseSensitive.SetExStyle( WS_EX_TRANSPARENT, false );
    m_CaseSensitive.SetBackground( m_crBackground );
    m_CaseSensitive.Create( m_hWindow );
    m_CaseSensitive.GetSize( &size );
    m_CaseSensitive.SetSize( 175, size.cy );

    // Static label for text not found
    m_TextNotFound.SetBackground( m_crBackground );
    m_TextNotFound.Create( m_hWindow );
    m_TextNotFound.GetSize( &size );
    m_TextNotFound.SetSize( 200, size.cy );
    m_TextNotFound.SetVisible( false );

    m_bControlsCreated = true;
    DoLayout();
}

//===============================================================================================//
//  Description:
//      Set the input focus on the edit box
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void FindTextBar::FocusOnTextBox()
{
    if ( m_TextField.GetHwnd() )
    {
        SetFocus( m_TextField.GetHwnd() );
    }
}

//===============================================================================================//
//  Description:
//      Get the window handle of the backward find button
//
//  Parameters:
//      None
//
//  Returns:
//      HWND, NULL if the control has not been created
//===============================================================================================//
HWND FindTextBar::GetFindBackwardHwnd() const
{
    return m_FindBackward.GetHwnd();
}

//===============================================================================================//
//  Description:
//      Get the window handle of the forward find button
//
//  Parameters:
//      None
//
//  Returns:
//      HWND, NULL if the control has not been created
//===============================================================================================//
HWND FindTextBar::GetFindFowardHwnd() const
{
    return m_FindForward.GetHwnd();
}

//===============================================================================================//
//  Description:
//      Get this window's preffered height
//
//  Parameters:
//      None
//
//  Returns:
//      int pixel height
//===============================================================================================//
int FindTextBar::GetPreferredHeight() const
{
    return FIND_TEXT_BAR_HEIGHT;
}

//===============================================================================================//
//  Description:
//      Get the parameters for the search
//
//  Parameters:
//      pText          - string object to receive the text to search for
//      pCaseSensitive - flag to receive if search is to be case sensitive
//
//  Returns:
//      void
//===============================================================================================//
void FindTextBar::GetSearchParameters( String* pText, bool* pCaseSensitive ) const
{
    if ( ( pText == nullptr ) || ( pCaseSensitive == nullptr ) )
    {
        throw ParameterException( L"pText/pCaseSensitive", __FUNCTION__ );
    }
    m_TextField.GetText( pText );
    *pCaseSensitive = m_CaseSensitive.GetState();
}

//===============================================================================================//
//  Description:
//      Get the window handle of the text field box
//
//  Parameters:
//      None
//
//  Returns:
//      HWND, NULL if the control has not been created
//===============================================================================================//
HWND FindTextBar::GetTextFieldHwnd() const
{
    return m_TextField.GetHwnd();
}

//===============================================================================================//
//  Description:
//      Set the image button colours
//
//  Parameters:
//      buttonFillColour - the colour to filll the button's bitmap
//
//  Returns:
//      void
//===============================================================================================//
void FindTextBar::SetControlColours( COLORREF buttonFillColour )
{
    m_Find.SetBackground( m_crBackground );

    m_FindForward.SetFilledBitmap( buttonFillColour, PXS_COLOUR_BLACK, PXS_SHAPE_ARROW_DOWN );
    m_FindForward.SetBackground( m_crBackground );
    m_FindForward.SetBackgroundGradient( m_crGradient1, m_crGradient2, true );

    m_FindBackward.SetBackground( m_crBackground );
    m_FindBackward.SetFilledBitmap( buttonFillColour, PXS_COLOUR_BLACK, PXS_SHAPE_ARROW_UP );
    m_FindBackward.SetBackgroundGradient( m_crGradient1, m_crGradient2, true );

    m_CaseSensitive.SetBackground( m_crBackground );

    m_TextNotFound.SetBackground( m_crBackground );
}

//===============================================================================================//
//  Description:
//      Set the text for the controls on the find bar
//
//  Parameters:
//      Find          - caption for the "Find:" label
//      FindBackwards - caption for the "Find Backwards" tool tip
//      FindForwards  - caption for the "Find Forwards" tool tip
//      MatchCase     - caption for the Match Case checkbox
//      TextNotFound  - caption for the "Text Not Found" label
//
//  Remarks;
//      Must have created the controls before calling this function
//
//  Returns:
//      void
//===============================================================================================//
void FindTextBar::SetCaptions( const String& Find,
                               const String& FindBackwards,
                               const String& FindForwards,
                               const String& MatchCase, const String& TextNotFound )
{
    SIZE size = { 0, 0 };

    // Controls must have been created
    if ( m_bControlsCreated == false )
    {
        throw FunctionException( L"m_bControlsCreated", __FUNCTION__ );
    }

    m_Find.SetText( Find );
    m_FindBackward.SetToolTipText( FindBackwards );
    m_FindForward.SetToolTipText( FindForwards );
    m_CaseSensitive.SetText( MatchCase );
    m_TextNotFound.SetText ( TextNotFound );

    // Set sizes layout
    m_Find.GetPreferredSize( &size );
    m_Find.SetSize( size );

    memset( &size, 0, sizeof ( size ) );
    m_TextNotFound.GetPreferredSize( &size );
    m_TextNotFound.SetSize( size );

    Repaint();
}

//===============================================================================================//
//  Description:
//      Set the parameters for the search
//
//  Parameters:
//      Text          - string object to receive the text to search for
//      caseSensitive - flag to receive if search is to be case sensitive
//
//  Returns:
//      void
//===============================================================================================//
void FindTextBar::SetSearchParameters( const String& Text, bool caseSensitive )
{
    m_TextField.SetText( Text );
    m_CaseSensitive.SetState( caseSensitive );
}

//===============================================================================================//
//  Description:
//      Show/Hide the "Text not found message"
//
//  Parameters:
//      visible - true to show, false to hide
//
//  Returns:
//      void
//===============================================================================================//
void FindTextBar::ShowTextNotFoundLabel( bool visible )
{
    m_TextNotFound.SetVisible ( visible );
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//    Description:
//        Handle WM_COMMAND events.
//
//    Parameters:
//        wParam
//        lParam
//
//    Returns:
//        0 if handled, else non-zero.
//===============================================================================================//
LRESULT FindTextBar::CommandEvent( WPARAM wParam, LPARAM lParam )
{
    LRESULT result = 0;

    // Check that the controls have been created
    if ( m_bControlsCreated == false )
    {
        return 0;
    }

    // Pass commands to the event listener
    if ( m_hWndAppMessageListener )
    {
        // On EN_CHANGE, tell the event listener to find the text
        if ( ( HIWORD( wParam ) == EN_CHANGE ) &&
             ( m_TextField.GetHwnd() == reinterpret_cast<HWND>( lParam ) ) )
        {
            result = SendMessage( m_hWndAppMessageListener, WM_COMMAND, wParam, lParam );
        }
    }

    return result;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
