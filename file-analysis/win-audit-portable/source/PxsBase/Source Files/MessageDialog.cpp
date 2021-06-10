///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Message Dialog Class Implementation
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
// Include Files
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/MessageDialog.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/StaticControl.h"


#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/StringT.h"
///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
MessageDialog::MessageDialog()
              :m_bIsError( false ),
               m_Message(),
               m_Copy(),
               m_OK(),
               m_MenuPopup(),
               m_MenuSelectAll(),
               m_MenuCopy(),
               m_RichEditBox()
{
    try
    {
        SetSize( 400, 350 );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

// Copy constructor, not allowed so no implementation

// Destructor
MessageDialog::~MessageDialog()
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
//      Set flag to say this dialog is showing an error message
//
//  Parameters:
//      isError - true for error message
//
//  Returns:
//      void
//===============================================================================================//
void MessageDialog::SetIsError( bool isError )
{
    m_bIsError = isError;
}

//===============================================================================================//
//  Description:
//      Set the text to show on the dialogue box
//
//  Parameters:
//      Message - the message to display
//
//  Returns:
//      void
//===============================================================================================//
void MessageDialog::SetMessage( const String& Message )
{
    m_Message = Message;
    m_Message.Trim();
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Handle WM_COMMAND events.
//
//  Parameters:
//        wParam
//        lParam
//
//  Returns:
//        0 if handled, else non-zero.
//===============================================================================================//
LRESULT MessageDialog::CommandEvent( WPARAM wParam, LPARAM lParam )
{
    WORD    id     = LOWORD(wParam);    // Item, control, or accelerator
    LRESULT result = 0;                 // Handled

    // Ensure controls have been created
    if ( m_bControlsCreated == false )
    {
        return 1;   // Not handled
    }

    if ( ( id == IDCANCEL ) || IsClickFromButton( m_OK, wParam, lParam ) )
    {
        EndDialog( m_hWindow, IDCANCEL );
    }
    else if ( IsClickFromButton( m_Copy, wParam, lParam ) )
    {
        m_RichEditBox.SelectAll();
        m_RichEditBox.CopySelection();
        m_RichEditBox.RemoveSelection();
        m_RichEditBox.ScrollToTop();
        m_Copy.Repaint();
    }
    else if ( id == PXS_APP_MSG_SELECT_ALL )
    {
        m_RichEditBox.SelectAll();
    }
    else if ( id == PXS_APP_MSG_COPY_SELECTION )
    {
        m_RichEditBox.CopySelection();
    }
    else
    {
        result = 1;     // Not handled
    }

    return result;
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
void MessageDialog::NotifyEvent( const NMHDR* pNmhdr )
{
    const MSGFILTER* pMsgFilter = nullptr;

    if ( pNmhdr == nullptr )
    {
        return;     // Nothing to do
    }

    if ( pNmhdr->code == EN_MSGFILTER )
    {
        // Show the pop-up menu,  pNmhdr is first member of MSGFILTER
        pMsgFilter = reinterpret_cast<const MSGFILTER*>( pNmhdr );
        if ( pMsgFilter->msg == WM_RBUTTONUP )
        {
            bool  textSelected = m_RichEditBox.IsAnyTextSelected();
            POINT point     = { 0, 0 };
            HMENU hEditMenu = m_MenuPopup.GetMenuHandle();

            m_MenuCopy.SetEnabled( textSelected );
            if ( hEditMenu )
            {
                GetCursorPos( &point );
                TrackPopupMenu( hEditMenu, 0, point.x, point.y, 0, m_hWindow, nullptr );
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Handle the WM_INITDIALOG message
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void MessageDialog::InitDialogEvent()
{
    SIZE   buttonSize = { 0, 0 }, dialogSize = { 0, 0 };
    RECT   bounds     = { 0, 0, 0, 0 };
    POINT  location   = { 0, 0 };
    HICON  hIcon      = nullptr;
    Font   FontObject;
    String Text, ExePath;
    StaticControl Static;

    // Only allow initialization once
    if ( m_bControlsCreated )
    {
        return;
    }
    m_bControlsCreated = true;

    if ( m_hWindow == nullptr )
    {
        return;
    }
    GetClientSize( &dialogSize );

    // Set a font
    FontObject.SetFaceName( L"Verdana" );
    FontObject.SetPointSize( 9, nullptr );
    FontObject.Create();

    // Icon
    bounds.left    = 10;
    bounds.top     = dialogSize.cy - 10 - GetSystemMetrics( SM_CYICON );
    bounds.right   = bounds.left + 40;
    bounds.bottom  = bounds.top  + 40;
    Static.SetBounds( bounds );
    Static.SetAlignmentX( PXS_CENTER_ALIGNMENT );
    Static.SetAlignmentY( PXS_CENTER_ALIGNMENT );
    if ( m_bIsError )
    {
        hIcon = LoadIcon( nullptr, IDI_ERROR );
        Static.SetIcon( hIcon );
    }
    else
    {
        // Use the first icon in the exe file.
        hIcon = LoadIcon( GetModuleHandle( nullptr ), MAKEINTRESOURCE( 1 ) );
    }
    try
    {
        if ( hIcon )
        {
            Static.SetIcon( hIcon );
        }
    }
    catch ( const Exception& eIcon )
    {
        PXSLogException( eIcon, __FUNCTION__ );
    }
    m_Statics.Add( Static );
    Static.Reset();

    // Standard button size
    buttonSize.cx = 80;
    buttonSize.cy = 25;

    // Close button - add first so it has focus
    PXSGetResourceString( PXS_IDS_130_CLOSE, &Text );
    location.x = dialogSize.cx - ( 10 + buttonSize.cx );
    location.y = dialogSize.cy - ( 10 + buttonSize.cy );
    m_OK.SetLocation( location );
    m_OK.Create( m_hWindow );
    m_OK.SetText( Text );

    // Create the edit box
    bounds.left   = 10;
    bounds.top    = 10;
    bounds.bottom = location.y - 20;
    bounds.right  = dialogSize.cx - 10;
    m_RichEditBox.SetBounds( bounds );
    m_RichEditBox.SetReadOnly( true );
    m_RichEditBox.WrapText( true );
    m_RichEditBox.Create( m_hWindow );
    m_RichEditBox.SetFont( FontObject );
    m_RichEditBox.SetMargins( 10, 10 );
    m_RichEditBox.AddToEventMask( ENM_MOUSEEVENTS );
    m_RichEditBox.SetRichText( m_Message );

    // Horizontal line
    bounds.left   = 10;
    bounds.right  = dialogSize.cx - 10;
    bounds.top    = location.y - 10;
    bounds.bottom = bounds.top + 2;
    Static.SetBounds( bounds );
    Static.SetShape( PXS_SHAPE_SUNK );
    m_Statics.Add( Static );
    Static.Reset();

    location.y = dialogSize.cy - ( 10 + buttonSize.cy );

    // Copy button
    PXSGetResourceString( PXS_IDS_129_COPY, &Text );
    location.x = dialogSize.cx - ( 2 * ( 10 + buttonSize.cx ) );
    m_Copy.SetLocation( location );
    m_Copy.Create( m_hWindow );
    m_Copy.SetText( Text );

    // Pop-up menu
    m_MenuPopup.Show( nullptr, 0 );       // nullptr = floating

    PXSGetResourceString( PXS_IDS_131_SELECT_ALL, &Text );
    m_MenuSelectAll.SetLabel( Text );
    m_MenuSelectAll.SetCommandID( PXS_APP_MSG_SELECT_ALL );
    m_MenuSelectAll.SetFilledBitmap( PXS_COLOUR_TEAL, PXS_COLOUR_BLACK );

    PXSGetResourceString( PXS_IDS_129_COPY, &Text );
    m_MenuCopy.SetLabel( Text );
    m_MenuCopy.SetCommandID( PXS_APP_MSG_COPY_SELECTION );
    m_MenuCopy.SetFilledBitmap( PXS_COLOUR_GREEN, PXS_COLOUR_BLACK );

    m_MenuPopup.AddMenuItem( &m_MenuSelectAll );
    m_MenuPopup.AddSeparator();
    m_MenuPopup.AddMenuItem( &m_MenuCopy );

    // Mirror for RTL
    if ( IsRightToLeftReading() )
    {
        RtlStatics();
        m_Copy.RtlMirror( dialogSize.cx );
        m_OK.RtlMirror( dialogSize.cx );
        m_RichEditBox.RtlMirror( dialogSize.cx );
    }
    m_bControlsCreated = true;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////
