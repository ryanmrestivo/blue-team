///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Tree View Class Implementation
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
#include "PxsBase/Header Files/TreeView.h"

// 2. C System Files
#include <ShlObj.h>

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/AllocateBytes.h"
#include "PxsBase/Header Files/Application.h"
#include "PxsBase/Header Files/Directory.h"
#include "PxsBase/Header Files/Exception.h"
#include "PxsBase/Header Files/File.h"
#include "PxsBase/Header Files/Formatter.h"
#include "PxsBase/Header Files/ImageButton.h"
#include "PxsBase/Header Files/MemoryException.h"
#include "PxsBase/Header Files/MenuItem.h"
#include "PxsBase/Header Files/MenuPopup.h"
#include "PxsBase/Header Files/ParameterException.h"
#include "PxsBase/Header Files/StaticControl.h"
#include "PxsBase/Header Files/TreeViewItem.h"
#include "PxsBase/Header Files/UTF8.h"
#include "PxsBase/Header Files/WaitCursor.h"

///////////////////////////////////////////////////////////////////////////////////////////////////
// Construction/Destruction
///////////////////////////////////////////////////////////////////////////////////////////////////

// Default constructor
TreeView::TreeView()
         :NODE_SQUARE_SIZE( 9 ),
          m_nIdxSelectedFirst( -1 ),
          m_nIdxSelectedLast( -1 ),
          m_nHiLiteLine( -1 ),
          m_uCollapseAllBitmapId( 0 ),
          m_uExpandAllBitmapId( 0 ),
          m_uSelectAllBitmapId( 0 ),
          m_uDeselectAllBitmapId( 0 ),
          m_crClientTitleBar( PXS_COLOUR_WHITE ),
          m_crTextHiLiteBackground( PXS_COLOUR_NAVY ),
          m_hBitmapNodeOpen( nullptr ),
          m_hBitmapNodeClosed( nullptr ),
          m_hBitmapLeaf( nullptr ),
          m_wIndentTab( 16 ),
          m_LastWindowRect(),        // Non-op
          m_LastHiLiteLineRect(),    // Non-op
          m_CollapseAll(),
          m_ExpandAll(),
          m_SelectAll(),
          m_DeselectAll(),
          m_ClientTitleBarText(),
          m_Items()
{
    memset( &m_LastWindowRect    , 0, sizeof ( m_LastWindowRect ) );
    memset( &m_LastHiLiteLineRect, 0, sizeof ( m_LastHiLiteLineRect ) );

    // Creation parameters. Want to clip children so that the close button
    // in the client title does not flicker when the window is repainted
    m_CreateStruct.style     |= WS_CLIPCHILDREN;
    m_CreateStruct.dwExStyle |= WS_EX_CLIENTEDGE;

    // Colours
    m_crClientTitleBar = GetSysColor( COLOR_BTNFACE );
}

// Copy constructor - not allowed so no implementation

// Destructor
TreeView::~TreeView()
{
    // Release bitmap resources
    if ( m_hBitmapNodeOpen )
    {
        DeleteObject( m_hBitmapNodeOpen );
    }

    if ( m_hBitmapNodeClosed )
    {
        DeleteObject( m_hBitmapNodeClosed );
    }

    if ( m_hBitmapLeaf )
    {
        DeleteObject( m_hBitmapLeaf );
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
//      Add an item to the list
//
//  Parameters:
//      TreeViewItem - a tree view item
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::AddItem( const TreeViewItem& TreeViewItem )
{
    SIZE preferredSize = {0, 0 };

    m_Items.Append( TreeViewItem );
    EvaluatePreferredSize( &preferredSize );
    UpdateScrollBarsInfo( preferredSize );
    Repaint();
}

//===============================================================================================//
//  Description:
//      Remove all of the items from the list
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::ClearList()
{
    SIZE preferredSize = { 0, 0 };

    m_Items.RemoveAll();

    // Set to not item selected
    m_nIdxSelectedFirst = -1;
    m_nIdxSelectedLast  = -1;
    EvaluatePreferredSize( &preferredSize );
    UpdateScrollBarsInfo( preferredSize );
    Repaint();
}

//===============================================================================================//
//  Description:
//      Expand or collapse all nodes
//
//  Parameters:
//      expand - true if to expand all or false to collapse
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::ExpandOrCollapseAll( bool expand )
{
    TreeViewItem* pItem = nullptr;

    if ( m_Items.IsEmpty() )
    {
        return;     // Nothing to do
    }

    m_Items.Rewind();
    do
    {
        pItem = m_Items.GetPointer();
        if ( pItem->IsNode() )
        {
            pItem->SetExpanded( expand );
        }
    } while ( m_Items.Advance() );

    // Simulate a size event as the contents may have changed
    memset( &m_LastWindowRect, 0, sizeof ( m_LastWindowRect ) );
    SizeEvent();
}

//===============================================================================================//
//  Description:
//      Set the background colour of the client area title bar
//
//  Parameters:
//      clientTitleBarColour - color
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::SetClientTitleBarColour( COLORREF clientTitleBarColour )
{
    m_crClientTitleBar = clientTitleBarColour;
}

//===============================================================================================//
//  Description:
//      Set the text of a title bar area in the panel
//
//  Parameters:
//      ClientTitleBarText - text for the client area title
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::SetClientTitleBarText( const String& ClientTitleBarText )
{
    m_ClientTitleBarText = ClientTitleBarText;
    Repaint();
}

//===============================================================================================//
//  Description:
//      Set the labels for the pop-up menu
//
//  Parameters:
//      CollapseAll - string for "Collapse All"
//      ExpandAll   - string for "Expand All"
//      SelectAll   - string for "Select All"
//      DeselectAll - string for "Deselect All"
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::SetMenuLabels( const String& CollapseAll,
                              const String& ExpandAll,
                              const String& SelectAll, const String& DeselectAll )
{
    m_CollapseAll  = CollapseAll;
    m_ExpandAll    = ExpandAll;
    m_SelectAll    = SelectAll;
    m_DeselectAll  = DeselectAll;
}

//===============================================================================================//
//  Description:
//      Select or deselect all items
//
//  Parameters:
//      selectAll - true if want all items selected, otherwise false
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::SetSelectAll( bool selectAll )
{
    TreeViewItem* pItem = nullptr;

    if ( m_Items.IsEmpty() )
    {
        return;     // Nothing to do
    }

    m_Items.Rewind();
    do
    {
        pItem = m_Items.GetPointer();
        pItem->SetSelected( selectAll );
    } while ( m_Items.Advance() );

    Repaint();
}

//===============================================================================================//
//  Description:
//      Set the bitmaps for the list
//
//  Parameters:
//      nodeClosedID        - the resource ID of the closed node bitmap
//      nodeOpenID          - resource ID of the open node bitmap
//      leafID              - the resource ID of the leaf bitmap
//      transparent         - transparent colour for the bitmaps
//      collapseAllBitmapId - resource ID for "Collapse All" on pop-up menu
//      expandAllBitmapId   - resource ID for "Expand All" on pop-up menu
//      selectAllBitmapId   - resource ID for "Select All" on pop-up menu
//      deselectAllBitmapId - resource ID for "Deselect All" on pop-up menu
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::SetBitmaps( WORD nodeClosedID,
                           WORD nodeOpenID,
                           WORD leafID,
                           COLORREF transparent,
                           WORD collapseAllBitmapId,
                           WORD expandAllBitmapId,
                           WORD selectAllBitmapId, WORD deselectAllBitmapId )
{
    // Release existing bitmap resources
    if ( m_hBitmapNodeOpen )
    {
        DeleteObject( m_hBitmapNodeOpen );
        m_hBitmapNodeOpen = nullptr;
    }

    if ( m_hBitmapNodeClosed )
    {
        DeleteObject( m_hBitmapNodeClosed );
        m_hBitmapNodeClosed = nullptr;
    }

    if ( m_hBitmapLeaf )
    {
        DeleteObject( m_hBitmapLeaf );
        m_hBitmapLeaf = nullptr;
    }

    // Load the bitmaps, and make them transparent
    if ( nodeClosedID )
    {
        m_hBitmapNodeClosed = LoadBitmap( GetModuleHandle( nullptr ),
                                          MAKEINTRESOURCE( nodeClosedID ) );
    }

    if ( nodeOpenID )
    {
        m_hBitmapNodeOpen = LoadBitmap( GetModuleHandle( nullptr ),
                                        MAKEINTRESOURCE( nodeOpenID ) );
    }

    if ( leafID )
    {
        m_hBitmapLeaf = LoadBitmap( GetModuleHandle( nullptr ),
                                    MAKEINTRESOURCE( leafID ) );
    }

    // Make them transparent
    if ( ( transparent == CLR_INVALID ) && ( m_crBackground != CLR_INVALID ) )
    {
        PXSReplaceBitmapColour( m_hBitmapNodeClosed, transparent, m_crBackground );
        PXSReplaceBitmapColour( m_hBitmapNodeOpen  , transparent, m_crBackground );
        PXSReplaceBitmapColour( m_hBitmapLeaf      , transparent, m_crBackground );
    }
    Repaint();

    // Store these string resource identifiers
    m_uCollapseAllBitmapId = collapseAllBitmapId;
    m_uExpandAllBitmapId   = expandAllBitmapId;
    m_uSelectAllBitmapId   = selectAllBitmapId;
    m_uDeselectAllBitmapId = deselectAllBitmapId;
}

//===============================================================================================//
//  Description:
//      Set the text highlight colour
//
//  Parameters:
//      textHiLiteBackground - the high light colour
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::SetTextHiLiteBackground( COLORREF textHiLiteBackground )
{
    m_crTextHiLiteBackground = textHiLiteBackground;
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Protected Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Handle a WM_COMMAND event
//
//  Parameters:
//      wParam - The low-order word of wParam identifies the action code.
//               The high-order word is 0.
//
//      lParam - If message source was a window lParam is the HWND of
//               window, otherwise, lParam is 0.
//  Returns:
//      LRESULT of zero which implies handled
//===============================================================================================//
LRESULT TreeView::CommandEvent( WPARAM wParam, LPARAM /* lParam */ )
{
    WORD action = LOWORD(wParam);  // Item, control, or accelerator identifier

    if ( action == PXS_APP_MSG_COLLAPSE_ALL )
    {
        ExpandOrCollapseAll( false );
    }
    else if ( action == PXS_APP_MSG_EXPAND_ALL )
    {
        ExpandOrCollapseAll( true );
    }
    else if ( action == PXS_APP_MSG_SELECT_ALL )
    {
        SetSelectAll( true );
    }
    else if ( action == PXS_APP_MSG_DESELECT_ALL )
    {
        SetSelectAll( false );
    }

    return 0;       // Handled
}

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
LRESULT TreeView::KeyDownEvent( WPARAM virtKey )
{
    bool  processEvent  = true;
    int   lineNumber    = 0, numClientLines, numVisibleLines;
    WPARAM  keyState    = 0;
    TreeViewItem* pItem = nullptr;

    numClientLines  = GetNumClientLines();
    numVisibleLines = GetNumVisibleLines();

    // Get the state of the shift and control keys
    if ( GetKeyState( VK_SHIFT ) & 0x8000 )
    {
        keyState |= MK_SHIFT;
    }

    if ( GetKeyState( VK_CONTROL ) & 0x8000 )
    {
        keyState |= MK_CONTROL;
    }

    switch ( virtKey )
    {
        default:
            processEvent = false;
            break;

        ///////////////////////////////////////////////////
        // Key strokes that correspond to vertical movement

        case VK_PRIOR:              // page up key
            lineNumber  = LineNumberFromIndex( m_nIdxSelectedLast );
            lineNumber -= numClientLines;
            lineNumber  = PXSMaxInt( lineNumber, 1 );
            break;

        case VK_NEXT:               // page down key
            lineNumber  = LineNumberFromIndex( m_nIdxSelectedLast );
            lineNumber += PXSMinInt( numVisibleLines, numClientLines );
            break;

        case VK_END:                // end key
            lineNumber = GetNumVisibleLines();
            break;

        case VK_HOME:               // home key
            lineNumber = 1;
            break;

        case VK_UP:                 // up arrow key
            lineNumber = LineNumberFromIndex( m_nIdxSelectedLast );
            lineNumber--;
            PXSMaxInt( 1, lineNumber );
            break;

        case VK_DOWN:               // down arrow key
            lineNumber = LineNumberFromIndex( m_nIdxSelectedLast );
            lineNumber = PXSMinInt( numVisibleLines, lineNumber + 1 );
            break;

        ///////////////////////////////////////////////////////////////////////
        // Key strokes that correspond to horizontal movement.
        // Although there is no vertical movement, the size of the contents
        // can change so will process a size event.

        case VK_LEFT:               // left arrow key
        case VK_RIGHT:              // right arrow key

            // If on a node, collapse it
            processEvent = false;
            pItem = GetLastSelectedItem();
            if ( pItem && pItem->IsNode() )
            {
                if ( virtKey == VK_RIGHT )
                {
                    pItem->SetExpanded( true );
                }
                else
                {
                    pItem->SetExpanded( false );
                }
            }
            memset( &m_LastWindowRect, 0, sizeof ( m_LastWindowRect ) );
            SizeEvent();
            break;
    }

    // If required, process the key stroke
    if ( processEvent )
    {
        // Process the event, set flag to says its a keyboard event
        HandleLineSelectionEvent( lineNumber, keyState, false );
        UpdateScrollPosition( lineNumber );
        InformAppMessageListener();
    }

    return 0;   // Say always handled
}

//===============================================================================================//
//  Description:
//      Handle WM_LBUTTONDOWN event
//
//  Parameters:
//      point - point in window where the left button of mouse was clicked
//      keys  - which virtual keys are down
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::MouseLButtonDownEvent( const POINT& point, WPARAM keys )
{
    SIZE preferredSize = { 0, 0 };

    int lineNumber;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    // If clicked on the title bar area then nothing to do
    lineNumber = MouseOverLineNumber( point );
    if ( lineNumber == 0 )
    {
        return;
    }

    if ( m_hWindow != GetFocus() )
    {
        SetFocus( m_hWindow );
    }
    WaitCursor Wait;
    HandleLineSelectionEvent( lineNumber, keys, true );
    InformAppMessageListener();
    EvaluatePreferredSize( &preferredSize );
    UpdateScrollBarsInfo( preferredSize );
    Repaint();
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
void TreeView::MouseMoveEvent( const POINT& point, WPARAM /* keys */ )
{
    RECT  clientRect     = { 0, 0, 0, 0 };
    POINT scrollPosition = { 0, 0 };

    if ( m_hWindow == nullptr )
    {
        return;
    }
    GetClientRect( m_hWindow, &clientRect );
    SetCursor( LoadCursor( nullptr, IDC_ARROW ) );

    // If this window is not contained in the foreground window
    // then ignore rest of processing
    if ( IsChild( GetForegroundWindow(), m_hWindow ) == 0 )
    {
        return;
    }

    // Redraw the line the mouse was previously over, use the current width
    // of the window in case it has changed, since the last mouse movement.
    m_LastHiLiteLineRect.right = clientRect.right;
    InvalidateRect( m_hWindow, &m_LastHiLiteLineRect, FALSE );

    // Redraw the line the mouse is over, do not do the first one as this
    //  is the client's title bar
    m_nHiLiteLine = MouseOverLineNumber( point );
    if ( m_nHiLiteLine > 0 )
    {
        GetScrollPosition( &scrollPosition );
        clientRect.top   -= scrollPosition.y;
        clientRect.top   += ( m_nHiLiteLine - 1 ) * m_nScreenLineHeight;
        clientRect.bottom = clientRect.top + ( 2 * m_nScreenLineHeight );
        InvalidateRect( m_hWindow, &clientRect, FALSE );

        // Save rectangle for next pass
        CopyRect( &m_LastHiLiteLineRect, &clientRect );
    }

    // Handle redraw the close button if mouse is inside the bitmap
    if ( m_hBitmapHide )
    {
        if ( PtInRect( &m_recHideBitmap, point ) )
        {
            SetTimer( m_hWindow, HIDE_WINDOW_TIMER_ID, 100, nullptr );
        }
        InvalidateRect( m_hWindow, &m_recHideBitmap, FALSE );
    }
}

//===============================================================================================//
//
//  Description:
//        Handle the WM_RBUTTONDOWN event
//
//  Parameters:
//      point - POINT where mouse was right clicked in the client area
//
//  Remarks:
//      Select the item right-clicked on, this way when the user
//      releases the button any pop-up menu will seem to be associated
//      with a selected item
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::MouseRButtonDownEvent( const POINT& point )
{
    int lineNumber = 0, index = 0;
    TreeViewItem* pItem;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    // If clicked on the title bar area then nothing to do
    lineNumber = MouseOverLineNumber( point );
    if ( lineNumber == 0 )
    {
        return;
    }

    // Set the focus to this window
    if ( m_hWindow != GetFocus() )
    {
        SetFocus( m_hWindow );
    }

    // Get the item, if the item is already selected then do nothing,
    // otherwise de-select all and select this one
    pItem = GetItemAtLineNumber( lineNumber, &index );
    if ( pItem )
    {
        if ( pItem->IsSelected() == false )
        {
            SetSelectAll( false );
            pItem->SetSelected( true );
        }
    }
    Repaint();
}

//===============================================================================================//
//  Description:
//  Handle the WM_RBUTTONUP event
//
//  Parameters:
//      point - POINT where mouse was right clicked in the client area

//  Remarks:
//      Right click will present a menu, for this to be displayed
//      the window must have an action listener specified that will
//      process the commands sent from the menu.
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::MouseRButtonUpEvent( const POINT& point )
{
    bool  enabled = false;
    int   index, lineNumber;
    POINT screenPoint = point;
    const TreeViewItem* pItem;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    // Get the line number and item the mouse is over
    index      = -1;
    lineNumber = MouseOverLineNumber( point );
    pItem      = GetItemAtLineNumber( lineNumber, &index );

    // If clicked on the title bar or there is no item then show a simple menu
    if ( ( lineNumber == 0 ) || ( pItem == nullptr ) )
    {
        // Collapse all
        MenuItem mnuCollapseAll;
        mnuCollapseAll.SetLabel( m_CollapseAll );
        mnuCollapseAll.SetCommandID( PXS_APP_MSG_COLLAPSE_ALL );
        mnuCollapseAll.SetBitmap( m_uCollapseAllBitmapId );

        // Expand all
        MenuItem mnuExpandAll;
        mnuExpandAll.SetLabel( m_ExpandAll );
        mnuExpandAll.SetCommandID( PXS_APP_MSG_EXPAND_ALL );
        mnuExpandAll.SetBitmap( m_uExpandAllBitmapId );

        // Select All
        MenuItem mnuSelectAll;
        mnuSelectAll.SetLabel( m_SelectAll );
        mnuSelectAll.SetCommandID( PXS_APP_MSG_SELECT_ALL );
        mnuSelectAll.SetBitmap( m_uSelectAllBitmapId );

        // Deselect all
        MenuItem mnuDeselectAll;
        mnuDeselectAll.SetLabel( m_DeselectAll );
        mnuDeselectAll.SetCommandID( PXS_APP_MSG_DESELECT_ALL );
        mnuDeselectAll.SetBitmap( m_uDeselectAllBitmapId );

        // Create pop-up menu
        MenuPopup mnuPopup;
        mnuPopup.Show( nullptr, 0 );

        // Add items to pop up menu
        mnuPopup.AddMenuItem( &mnuCollapseAll);
        mnuPopup.AddMenuItem( &mnuExpandAll );
        mnuPopup.AddMenuItem( &mnuSelectAll );
        mnuPopup.AddMenuItem( &mnuDeselectAll );

        if ( m_Items.IsEmpty() == false )
        {
            enabled = true;     // There is at least one item
        }
        mnuExpandAll.SetEnabled( enabled );
        mnuCollapseAll.SetEnabled( enabled );
        mnuSelectAll.SetEnabled( enabled );

        // Show the menu
        ClientToScreen( m_hWindow, &screenPoint );
        TrackPopupMenu( mnuPopup.GetMenuHandle(),
                        TPM_LEFTALIGN | TPM_RIGHTBUTTON,
                        screenPoint.x, screenPoint.y, 0, m_hWindow, nullptr );
    }
    else
    {
        // Show a pop-up menu if have n message listener. Make sure clicked on
        // an item. LPARAM is the index of the item for which to show the menu
        if ( m_hWndAppMessageListener )
        {
            if ( ( index != -1 ) && pItem )
            {
                SendMessage( m_hWndAppMessageListener, PXS_APP_MSG_SHOW_MENU, 0, (LPARAM)index );
            }
        }
    }
}

//===============================================================================================//
//  Description:
//      Handle WM_PAINT event.
//
//  Parameters:
//      hdc - Handle to the device context to paint on
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::PaintEvent( HDC hdc )
{
    bool  rtlReading = false;
    int   xPos = 0, yPos = 0, lineNumber = 0;
    BYTE  currentIndent  = 0;
    HGDIOBJ oldFont      = nullptr;
    RECT  clientRect     = { 0, 0, 0, 0 };
    POINT scrollPosition = { 0, 0 };
    StaticControl Static;
    const TreeViewItem* pItem;

    if ( hdc == nullptr )
    {
        return;
    }
    DrawBackground( hdc );

    if ( m_Items.IsEmpty() )
    {
        DrawClientAreaTitleBar( hdc );
        DrawHideBitmap( hdc );
        return;
    }

    GetClientRect( m_hWindow, &clientRect );
    rtlReading = IsRightToLeftReading();
    GetScrollPosition( &scrollPosition );
    xPos = -scrollPosition.x;
    yPos = -scrollPosition.y;

    oldFont = SelectObject( hdc, GetStockObject( DEFAULT_GUI_FONT ) );
    lineNumber = 1;        // Line numbering starts at 1
    m_Items.Rewind();
    do
    {
        xPos  = -scrollPosition.x;
        pItem = m_Items.GetPointer();
        if ( ShouldShowItem( pItem, &currentIndent ) )
        {
            yPos += m_nScreenLineHeight;

            // Only need to draw if visible on screen, be conservative
            // and draw one extra line above top of window
            if ( ( yPos >= -m_nScreenLineHeight ) &&
                    ( yPos <= clientRect.bottom    )  )
            {
                DrawItemBackground( hdc,
                                    lineNumber, yPos, clientRect.right );
                xPos = DrawItemSquare( hdc, *pItem, xPos, yPos, rtlReading);
                xPos = DrawItemBitmap( hdc, *pItem, xPos, yPos, rtlReading);
                xPos = DrawItemLabel(  hdc, *pItem, xPos, yPos, rtlReading);
            }
            lineNumber++;
        }
    } while ( m_Items.Advance() );

    DrawClientAreaTitleBar( hdc );
    DrawHideBitmap( hdc );

    // Clean up
    if ( oldFont )
    {
        SelectObject( hdc, oldFont );
    }
}

//===============================================================================================//
//  Description:
//      Handle WM_SIZE event.
//
//  Parameters:
//      none
//
//  Returns:
//      void.
//===============================================================================================//
void TreeView::SizeEvent()
{
    RECT windowRect    = { 0, 0, 0, 0};
    SIZE preferredSize = { 0, 0 };

    if ( m_hWindow == nullptr )
    {
        return;
    }
    GetWindowRect( m_hWindow, &windowRect );

    if ( EqualRect( &windowRect, &m_LastWindowRect ) )
    {
        return;     // Nothing to do so return
    }
    CopyRect( &m_LastWindowRect, &windowRect );
    EvaluatePreferredSize( &preferredSize );
    UpdateScrollBarsInfo( preferredSize );
    PositionHideBitmap();
    Repaint();
}

///////////////////////////////////////////////////////////////////////////////////////////////////
// Private Methods
///////////////////////////////////////////////////////////////////////////////////////////////////

//===============================================================================================//
//  Description:
//      Draw the window's title bar, if it has one
//
//  Parameters:
//      hdc - the device context
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::DrawClientAreaTitleBar( HDC hdc )
{
    Font TextFont;
    RECT titleBounds = { 0, 0, 0, 0 };
    StaticControl Static;

    if ( ( hdc == nullptr ) || ( m_hWindow == nullptr ) )
    {
        return;
    }

    try     // Catch exceptions in case of repetitive error
    {
        // Set its bounds, use a line height of 17 pixels
        GetClientRect( m_hWindow, &titleBounds );
        titleBounds.bottom = titleBounds.top + 17;
        Static.SetBounds( titleBounds );

        // Set the text
        TextFont.SetBold( true );
        TextFont.SetPointSize( 8, nullptr );
        TextFont.Create();
        Static.SetFont( TextFont );
        Static.SetAlignmentY( PXS_TOP_ALIGNMENT );
        Static.SetText( m_ClientTitleBarText );

        // Set text alignment accounting for Right To Left Reading
        if ( IsRightToLeftReading() )
        {
            Static.SetAlignmentX( PXS_RIGHT_ALIGNMENT );
        }
        else
        {
            Static.SetAlignmentX( PXS_LEFT_ALIGNMENT );
        }

        Static.SetBackground( m_crClientTitleBar );
        Static.SetForeground( m_crForeground );
        Static.Draw( hdc );
    }
    catch ( const Exception& e )
    {
        PXSLogException( e, __FUNCTION__ );
    }
}

//===============================================================================================//
//  Description:
//      Draw the specified tree item's background and border
//
//  Parameters:
//      hdc        - the device context
//      lineNumber - the items line number
//      yPos       - the y-position of the item
//      clientWidth- the client width of the tree view
//
//  Returns:
//      void
//===============================================================================================//
void  TreeView::DrawItemBackground( HDC hdc, int lineNumber, int yPos, int clientWidth )
{
    RECT bounds = { 0, 0, 0, 0 };
    StaticControl Static;

    if ( ( hdc == nullptr ) )
    {
        return;     // Nothing to do
    }

    // Only need to draw if its the high-lighted item
    if ( m_nHiLiteLine == lineNumber )
    {
        // Set the bounds
        bounds.left   = 1;
        bounds.top    = yPos;
        bounds.right  = clientWidth - 1;
        bounds.bottom = bounds.top + m_nScreenLineHeight;
        Static.SetBounds( bounds );
        Static.SetBackgroundGradient( m_crGradient1, m_crGradient2, true );
        Static.SetShape( PXS_SHAPE_RECTANGLE_DOTTED );
        Static.SetShapeColour( m_crForeground );
        Static.Draw( hdc );
    }
}

//===============================================================================================//
//  Description:
//      Draw the specified item's bitmap
//
//  Parameters:
//      hdc        - the device context
//      Item       - the tree view item
//      xPos       - the x-position of the item
//      yPos       - the y-position of the item
//      rtlReading - indicates if the window is right to left reading
//
//  Returns:
//      the updated x-position
//===============================================================================================//
int TreeView::DrawItemBitmap( HDC hdc,
                              const TreeViewItem& Item, int xPos, int yPos, bool rtlReading )
{
    const   int PADDING = 5;
    BITMAP  bitmap;
    HBITMAP hBitmap = nullptr;

    if ( hdc == nullptr )
    {
        return xPos;     // Nothing to do
    }

    // Determine the bitmap
    hBitmap = Item.GetBitmap();
    if ( hBitmap == nullptr )
    {
        hBitmap = m_hBitmapLeaf;
        if ( Item.IsNode() )
        {
            hBitmap = m_hBitmapNodeClosed;
            if ( Item.IsExpanded() )
            {
                hBitmap = m_hBitmapNodeOpen;
            }
        }
    }

    if ( hBitmap )
    {
        memset( &bitmap, 0, sizeof ( bitmap ) );
        GetObject( hBitmap, sizeof ( bitmap), &bitmap );
        if ( rtlReading )
        {
            xPos -= bitmap.bmWidth;
        }
        PXSDrawTransparentBitmap( hdc, hBitmap, xPos, yPos + 1, m_crBackground );

        // Update the x-position
        if ( rtlReading )
        {
            xPos -= PADDING;
        }
        else
        {
            xPos += ( bitmap.bmWidth + PADDING );
        }
    }

    return xPos;
}

//===============================================================================================//
//  Description:
//      Draw the specified item's text caption
//
//  Parameters:
//      hdc        - the device context
//      Item       - the tree view item
//      xPos       - the x-position of the item
//      yPos       - the y-position of the item
//      rtlReading - indicates if the window is right to left reading
//
//  Returns:
//      the updated x-position
//===============================================================================================//
int TreeView::DrawItemLabel( HDC hdc,
                             const TreeViewItem& Item, int xPos, int yPos, bool rtlReading )
{
    const int PADDING  = 5;
    int       nSavedDC = 0, charLength;
    SIZE      textSize = { 0, 0 };
    LPCWSTR   pszLabel;

    if ( hdc == nullptr )
    {
        return xPos;     // Nothing to do
    }

    pszLabel = Item.GetLabel();
    if ( pszLabel == nullptr )
    {
        return xPos;
    }
    nSavedDC = SaveDC( hdc );

    // Draw hi-lighted if selected
    SetBkMode( hdc, TRANSPARENT );
    if ( Item.IsSelected() )
    {
        // Draw selected item in hi-lite
        SetBkMode( hdc, OPAQUE );
        SetTextColor( hdc, PXS_COLOUR_WHITE );
        if ( m_crTextHiLiteBackground != CLR_INVALID )
        {
            SetBkColor( hdc, m_crTextHiLiteBackground );
        }
    }

    charLength = lstrlen( pszLabel );
    GetTextExtentPoint32( hdc, pszLabel, charLength, &textSize );
    if ( rtlReading )
    {
        xPos -= textSize.cx;
        TextOut( hdc, xPos, yPos + 1, pszLabel, charLength );
        xPos -= PADDING;                    // Update the x-position
    }
    else
    {
        TextOut( hdc, xPos, yPos + 1, pszLabel, charLength );
        xPos += ( textSize.cx + PADDING );  // Update the x-position
    }

    // Reset
    RestoreDC( hdc, nSavedDC );

    return xPos;
}

//===============================================================================================//
//  Description:
//      Draw the square associated with an item
//
//  Parameters:
//      hdc        - the device context
//      Item       - the tree view item
//      xPos       - the x-position of the item
//      yPos       - the y-position of the item
//      rtlReading - indicates if the window is right to left reading
//
//  Returns:
//      the updated x-position
//===============================================================================================//
int TreeView::DrawItemSquare( HDC hdc,
                              const TreeViewItem& Item, int xPos, int yPos, bool rtlReading )
{
    const int PADDING = 10;
    int     x1 = 0, x2 = 0, y1 = 0, y2 = 0, width = 0;
    RECT    clientRect = { 0, 0, 0, 0 };
    HGDIOBJ oldPen;
    SCROLLINFO  si;

    if ( hdc == nullptr )
    {
        return xPos;     // Nothing to do
    }

    // Set the x-position
    if ( rtlReading )
    {
        // Get the available width
        GetClientRect( m_hWindow, &clientRect );
        memset( &si, 0, sizeof ( si ) );
        si.cbSize = sizeof ( si );
        si.fMask  = SIF_ALL;
        GetScrollInfo( m_hWindow, SB_HORZ, &si );
        width = PXSMaxInt( si.nMax - si.nMin, clientRect.right - clientRect.left );

        xPos += ( width
                  - ( Item.GetIndent() * m_wIndentTab )
                  - NODE_SQUARE_SIZE - 1 );
    }
    else
    {
        xPos += ( Item.GetIndent() * m_wIndentTab );
    }

    // Only nodes have squares
    if ( Item.IsNode() )
    {
        oldPen = SelectObject( hdc, GetStockObject( BLACK_PEN ) );
        if ( rtlReading )
        {
            x1 = xPos - ( NODE_SQUARE_SIZE / 2 );
        }
        else
        {
            x1 = xPos + ( NODE_SQUARE_SIZE / 2 );
        }
        x2 = x1 + NODE_SQUARE_SIZE - 1;
        y1 = yPos + ( m_nScreenLineHeight - NODE_SQUARE_SIZE ) / 2;
        y2 = y1 + NODE_SQUARE_SIZE - 1;

        MoveToEx( hdc, x1, y1, nullptr );
        LineTo(   hdc, x1, y2 );
        LineTo(   hdc, x2, y2 );
        LineTo(   hdc, x2, y1 );
        LineTo(   hdc, x1, y1 );

        // Draw the horizontal line in the square
        MoveToEx( hdc, x1 + 2, ( y1 + y2 ) / 2, nullptr );
        LineTo(   hdc, x2 - 1, ( y1 + y2 ) / 2 );

        // If this node is closed draw the vertical line
        if ( Item.IsExpanded() == false )
        {
            MoveToEx( hdc, ( x1 + x2 ) / 2, y1 + 2, nullptr );
            LineTo(   hdc, ( x1 + x2 ) / 2, y2 - 1 );
        }

        // Reset
        if ( oldPen )
        {
            SelectObject( hdc, oldPen );
        }
    }

    // Update the x-position
    if ( rtlReading )
    {
        xPos -= PADDING;
    }
    else
    {
        xPos += ( NODE_SQUARE_SIZE + PADDING );
    }

    return xPos;
}

//===============================================================================================//
//  Description:
//      Evaluate the preferred size of this window
//
//  Parameters:
//      pPreferredSize - receives the size
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::EvaluatePreferredSize( SIZE* pPreferredSize )
{
    int     numLines    = 0;
    HDC     hDC         = nullptr;
    SIZE    labelSize = { 0, 0 }, clientSize = { 0, 0 };
    LONG    maxWidth  = 0, maxHeight = 0, itemWidth = 0;
    HGDIOBJ stockFont = nullptr, oldFont = nullptr;
    LPCWSTR pszLabel  = nullptr;
    const TreeViewItem* pItem = nullptr;

    // Set default size of 100 x 100
    if ( pPreferredSize == nullptr )
    {
        throw ParameterException( L"pPreferredSize", __FUNCTION__ );
    }
    pPreferredSize->cx = 100;
    pPreferredSize->cy = 100;

    if ( m_hWindow == nullptr )
    {
        return;
    }

    // Set the font
    hDC = GetDC( m_hWindow );
    if ( hDC == nullptr )
    {
        return;
    }
    stockFont = GetStockObject( DEFAULT_GUI_FONT );
    if ( stockFont )
    {
        oldFont = SelectObject( hDC, stockFont );
    }

    // Evaluate the maximum width
    try
    {
        if ( m_Items.IsEmpty() == false )
        {
            m_Items.Rewind();
            do
            {
                pItem = m_Items.GetPointer();
                // Get items width, stating with the indent and a possible
                // bitmap, space and gap as used when its drawn
                itemWidth  = pItem->GetIndent() * m_wIndentTab;
                itemWidth += m_wIndentTab;         // Node
                itemWidth += (16 + 2);             // Bitmap + spacer
                itemWidth +=  5;                   // Spacer

                memset( &labelSize, 0, sizeof ( labelSize ) );
                pszLabel = pItem->GetLabel();
                if ( hDC && pszLabel )
                {
                    GetTextExtentPoint32( hDC, pszLabel, lstrlen( pszLabel ), &labelSize );
                }
                itemWidth += labelSize.cx;
                maxWidth = PXSMaxInt( maxWidth, itemWidth );
            } while ( m_Items.Advance() );
        }
    }
    catch ( const Exception& )
    {
        if ( oldFont ) SelectObject( hDC, oldFont );
        ReleaseDC( m_hWindow, hDC );
        throw;
    }

    if ( oldFont )
    {
        SelectObject( hDC, oldFont );
    }
    ReleaseDC( m_hWindow, hDC );

    // Add 10 % more for good measure and width of possible vertical
    // scroll bar then take maximum of default and evaluated widths
    maxWidth += ( maxWidth / 10 );
    maxWidth += GetSystemMetrics( SM_CXVSCROLL );
    pPreferredSize->cx = PXSMaxInt( pPreferredSize->cx, maxWidth );

    // Adjust for the client area
    GetClientSize( &clientSize );
    pPreferredSize->cx -= clientSize.cx;
    pPreferredSize->cx  = PXSMaxInt( 0, pPreferredSize->cx );

    // Evaluate the height, account for possible horizontal scrollbar
    // then take maximum of default and evaluated heights
    numLines  = GetNumVisibleLines();
    maxHeight = numLines * m_nScreenLineHeight;

    // Add the height of 2 horizontal scroll bars, this
    // prevents resize cycling in the last line is partially obscured
    // and vertical scrollbar flips between visible/invisible
    maxHeight += GetSystemMetrics( SM_CYHSCROLL );
    maxHeight += GetSystemMetrics( SM_CYHSCROLL );
    pPreferredSize->cy = PXSMaxInt( pPreferredSize->cy, maxHeight );
}

//===============================================================================================//
//  Description:
//      Get a the index in the item lest at which the line corresponds
//
//  Parameters:
//      lineNumber - visible line
//      pIndex     - receives the index
//
//  Remarks:
//      Visible lines are all those that are visible by scrolling up
//      and down, not just those in the client area
//
//  Returns:
//      Pointer to the item, else NULL
//===============================================================================================//
TreeViewItem* TreeView::GetItemAtLineNumber( int lineNumber, int* pIndex )
{
    BYTE indent = 0;
    int  lineCounter = 0, indexCounter = 0;
    TreeViewItem* pItem = nullptr;

    if ( pIndex == nullptr )
    {
        throw ParameterException( L"pIndex", __FUNCTION__ );
    }
    *pIndex = -1;   // Reset

    if ( m_Items.IsEmpty() )
    {
        return nullptr;     // Nothing to do
    }

    lineCounter  = 0;
    indexCounter = 0;
    m_Items.Rewind();
    do
    {
        pItem = m_Items.GetPointer();
        if ( ( pItem->IsNode() ) && ( pItem->GetIndent() <= indent ) )
        {
            if ( pItem->IsExpanded() )
            {
                indent = (BYTE)( 0xFF & ( pItem->GetIndent() + 1 ) );
            }
            else
            {
                indent = pItem->GetIndent();
            }
        }

        // Count visible lines
        if ( pItem->GetIndent() <= indent )
        {
            lineCounter++;
        }

        // Check for a match
        if ( lineCounter == lineNumber )
        {
            *pIndex = indexCounter;
            break;
        }
        indexCounter++;
    } while ( m_Items.Advance() );

    return pItem;
}

//===============================================================================================//
//  Description:
//      Get a the number of visible lines
//
//  Parameters:
//      None
//
//  Remarks:
//      Visible lines are all those that are visible by scrolling up
//      and down, not just those in the client area
//
//  Returns:
//      int of number of visible lines
//===============================================================================================//
int TreeView::GetNumVisibleLines()
{
    BYTE  indent   = 0;
    int   numLines = 0;
    const TreeViewItem* pItem = nullptr;

    if ( m_Items.IsEmpty() )
    {
        return 0;
    }

    m_Items.Rewind();
    do
    {
        pItem = m_Items.GetPointer();
        // Determine the indent level
        if ( ( pItem->IsNode() ) && ( pItem->GetIndent() <= indent ) )
        {
            if ( pItem->IsExpanded() )
            {
                indent = (BYTE)( 0xFF & ( pItem->GetIndent() + 1 ) );
            }
            else
            {
                indent = pItem->GetIndent();
            }
        }

        // Count visible lines
        if ( pItem->GetIndent() <= indent )
        {
            numLines++;
        }
    } while ( m_Items.Advance() );
    m_Items.Rewind();

    return numLines;
}

//===============================================================================================//
//  Description:
//      Get a pointer to the selected tree view item
//
//  Parameters:
//      None
//
//  Returns:
//      pointer to tree view item, NULL if no item selected
//===============================================================================================//
TreeViewItem* TreeView::GetLastSelectedItem()
{
    int counter = 0;
    TreeViewItem* pItem  = nullptr;
    TreeViewItem* pFound = nullptr;

    // If no line selected, return NULL
    if ( m_nIdxSelectedLast == -1 )
    {
        return nullptr;
    }

    if ( m_Items.IsEmpty() )
    {
        return nullptr;
    }

    // Get matching item
    m_Items.Rewind();
    do
    {
        pItem = m_Items.GetPointer();
        if ( counter == m_nIdxSelectedLast )
        {
            pFound = pItem;
        }
        counter++;
    } while ( ( pFound == nullptr ) && m_Items.Advance() );

    return pFound;
}

//===============================================================================================//
//  Description:
//      Handle a selection event for a line
//
//  Parameters:
//      lineNumber - line number where the event occurred
//      keyState   - the key state
//      mouseEvent - flag to indicate if this is a mouse event
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::HandleLineSelectionEvent( int lineNumber,
                                         WPARAM keyState, bool mouseEvent )
{
    int index = 0;
    TreeViewItem* pItem = GetItemAtLineNumber( lineNumber, &index );

    if ( ( pItem == nullptr ) || ( index == -1 ) )
    {
        return;     // Nothing to do
    }

    // Set the selection range
    if ( m_nIdxSelectedFirst == -1 )
    {
        m_nIdxSelectedFirst = index;
    }
    m_nIdxSelectedLast = index;

    if ( mouseEvent && ( MK_CONTROL & keyState ) )
    {
        pItem->SetSelected( !pItem->IsSelected() );     // Toggle
    }
    else if ( MK_SHIFT & keyState )
    {
        SetSelectedItems();
    }
    else
    {
        // Clear entire selection, select this one and
        // set the last selected
        SetSelectAll( false );
        pItem->SetSelected( true );
        m_nIdxSelectedFirst = index;
        m_nIdxSelectedLast  = index;

        // If it is a node and this is a mouse event, toggle it
        if ( mouseEvent && pItem->IsNode() )
        {
            pItem->SetExpanded( !pItem->IsExpanded() );

            // The contents have changed so call the size
            // handler to refresh the scrollbars
            memset( &m_LastWindowRect, 0, sizeof ( m_LastWindowRect ) );
            SizeEvent();
        }
    }
}

//===============================================================================================//
//  Description:
//      Inform this window's message listener that item was selected
//
//  Parameters:
//      None
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::InformAppMessageListener()
{
    size_t numDataBytes;
    AllocateBytes   AllocBytes;
    COPYDATASTRUCT  copyData;
    const TreeViewItem* pItem;

    if ( m_hWindow && m_hWndAppMessageListener )
    {
        // Get the selected Item
        pItem = GetLastSelectedItem();
        if ( pItem && pItem->GetStringData() )
        {
            // Fill copy data block and use SendMessage as the data contains
            // pointers.
            numDataBytes = wcslen( pItem->GetStringData() );
            numDataBytes = PXSAddSizeT( numDataBytes, 1 );  // Terminator
            numDataBytes = PXSMultiplySizeT( numDataBytes, sizeof ( wchar_t ) );

            // Make a copy of data, as don't want to point into string object
            memset( &copyData, 0, sizeof ( copyData ) );
            copyData.lpData = AllocBytes.New( numDataBytes );
            copyData.dwData = 0;
            copyData.cbData = PXSCastSizeTToUInt32( numDataBytes );
            memcpy( copyData.lpData, pItem->GetStringData(), numDataBytes );

            SendMessage( m_hWndAppMessageListener,
                         WM_COPYDATA, (WPARAM)m_hWindow, (LPARAM)&copyData );
        }
    }
}

//===============================================================================================//
//  Description:
//      Get a the line number of the item at a given index
//
//  Parameters:
//  index - index of the item
//
//  Remarks:
//      Item must be visible
//
//  Returns:
//      int of line number, else -1
//===============================================================================================//
int TreeView::LineNumberFromIndex( int index )
{
    int   lineCounter = 0, indexCounter = 0, lineNumber = -1;
    BYTE  indent = 0;
    const TreeViewItem* pItem = nullptr;

    if ( m_Items.IsEmpty() )
    {
        return -1;
    }

    m_Items.Rewind();
    do
    {
        pItem = m_Items.GetPointer();
        if ( ( pItem->IsNode() ) && ( pItem->GetIndent() <= indent ) )
        {
            if ( pItem->IsExpanded() )
            {
                indent = (BYTE)( 0xFF & ( pItem->GetIndent() + 1 ) );
            }
            else
            {
                indent = pItem->GetIndent();
            }
        }

        // Count visible lines
        if ( pItem->GetIndent() <= indent )
        {
            lineCounter++;
        }

        if ( index == indexCounter )
        {
            lineNumber = lineCounter;
            break;
        }
        indexCounter++;
    } while ( m_Items.Advance() );

    return lineNumber;
}

//===============================================================================================//
//  Description:
//      Determine the line number the cursor is over
//
//  Parameters:
//      point - position of the cursor in the window's client area
//
//  Returns:
//      int of line number, -1 if unknown
//===============================================================================================//
int TreeView::MouseOverLineNumber( const POINT& point )
{
    int   lineNumber = -1;
    POINT scrollPosition = { 0, 0 };

    if ( m_hWindow == nullptr )
    {
        return lineNumber;
    }

    // Line numbering begins at 1
    if ( GetScrollPosition( &scrollPosition ) )
    {
        if ( point.y <= m_nScreenLineHeight )
        {
            lineNumber = 1;  // will decrement below
        }
        else if ( m_nScreenLineHeight > 0 )
        {
            lineNumber = 1 + ((scrollPosition.y + point.y)/m_nScreenLineHeight);
        }
    }
    else if ( m_nScreenLineHeight > 0 )
    {
        // On error, assume at scroll position zero
        lineNumber = 1 + ( point.y / m_nScreenLineHeight );
    }
    lineNumber -= 1;    // Decrement for client area title bar

    return lineNumber;
}

//===============================================================================================//
//  Description:
//      Set the state of the items that have been selected
//
//  Parameters:
//      none
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::SetSelectedItems()
{
    int i = 0;
    TreeViewItem* pItem = nullptr;

    if ( m_Items.IsEmpty() )
    {
        return;     // Nothing to do
    }

    // Scan list, select those in range.
    m_Items.Rewind();
    do
    {
        // Note, the selection maybe reversed
        pItem = m_Items.GetPointer();
        if ( ( (i >= m_nIdxSelectedFirst) && (i <= m_nIdxSelectedLast ) ) ||
             ( (i >= m_nIdxSelectedLast ) && (i <= m_nIdxSelectedFirst) )  )
        {
            pItem->SetSelected( true );
        }
        else
        {
            pItem->SetSelected( false );
        }
        i++;
    } while ( m_Items.Advance() );
}

//===============================================================================================//
//  Description:
//      Determine if should show the item based on if it is a node and expanded.
//
//  Parameters:
//      pItem          - the tree view item
//      pCurrentIndent - on input has the current indent level, on output
//                       receives the updated indent level
//
//  Returns:
//      true if the item should be drawn
//===============================================================================================//
bool TreeView::ShouldShowItem( const TreeViewItem* pItem, BYTE* pCurrentIndent )
{
    bool showItem = false;
    BYTE itemIndent;

    if ( ( pItem == nullptr ) || ( pCurrentIndent == nullptr ) )
    {
        return false;
    }

    itemIndent = pItem->GetIndent();
    if ( itemIndent <= *pCurrentIndent )
    {
        showItem = true;

        // Update the current indent level
        if ( pItem->IsNode() )
        {
            if ( pItem->IsExpanded() )
            {
                *pCurrentIndent = (BYTE)( 0xFF & ( itemIndent + 1 ) );
            }
            else
            {
                *pCurrentIndent = itemIndent;
            }
        }
    }

    return showItem;
}

//===============================================================================================//
//  Description:
//      Update the scroll bar position based on the specified line number
//
//  Parameters:
//      lineNumber - the line number
//
//  Returns:
//      void
//===============================================================================================//
void TreeView::UpdateScrollPosition( int lineNumber )
{
    SIZE  clientSize     = { 0, 0 };
    SIZE  preferedSized  = { 0, 0 };
    POINT newPosition    = { 0, 0 };
    POINT scrollPosition = { 0, 0 };

    GetClientSize( &clientSize );
    GetScrollPosition( &scrollPosition );

    // See if need to scroll the selected line into view
    newPosition.x = scrollPosition.x;
    if ( ( lineNumber * m_nScreenLineHeight ) <= scrollPosition.y )
    {
        // Move selected line to first line of the client area
        if ( lineNumber > 0 )
        {
            newPosition.y = ( ( lineNumber - 1 ) * m_nScreenLineHeight );
        }
        else
        {
            newPosition.y = 0;
        }
        SetScrollPosition( newPosition );
    }
    else if ( ( (lineNumber + 1) * m_nScreenLineHeight ) >=
                                          ( scrollPosition.y + clientSize.cy ) )
    {
        // Move selected line to last line of the client area
        newPosition.y = ( (lineNumber + 1)*m_nScreenLineHeight) - clientSize.cy;
        SetScrollPosition ( newPosition );
    }
    EvaluatePreferredSize( &preferedSized );
    UpdateScrollBarsInfo( preferedSized );
    Repaint();
}
