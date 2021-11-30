///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Tree View Class Header
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

#ifndef PXSBASE_TREE_VIEW_H_
#define PXSBASE_TREE_VIEW_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Remarks
///////////////////////////////////////////////////////////////////////////////////////////////////

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/PxsBase.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/ScrollPane.h"
#include "PxsBase/Header Files/TList.h"

// 6. Forwards
class TreeViewItem;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class TreeView : public ScrollPane
{
    public:
        // Default constructor
        TreeView();

        // Destructor
        ~TreeView();

        // Methods
        void    AddItem( const TreeViewItem& Item );
        void    ClearList();
        void    ExpandOrCollapseAll( bool expand );
        void    SetBitmaps( WORD nodeClosedID,
                            WORD nodeOpenID,
                            WORD leafID,
                            COLORREF  transparent,
                            WORD collapseAllBitmapId,
                            WORD expandAllBitmapId,
                            WORD selectAllBitmapId, WORD deselectAllBitmapId );
        void    SetClientTitleBarColour( COLORREF clientTitleBarColour );
        void    SetClientTitleBarText( const String& ClientTitleBarText );
        void    SetMenuLabels( const String& CollapseAll,
                               const String& ExpandAll,
                               const String& SelectAll, const String& DeselectAll );
        void    SetSelectAll( bool selectAll );
        void    SetTextHiLiteBackground( COLORREF textHiLiteBackground );

    protected:
        // Methods
        LRESULT CommandEvent( WPARAM wParam, LPARAM lParam );
        LRESULT KeyDownEvent( WPARAM virtKey );
        void    MouseLButtonDownEvent( const POINT& point, WPARAM keys );
        void    MouseMoveEvent( const POINT& point, WPARAM keys );
        void    MouseRButtonDownEvent( const POINT& point );
        void    MouseRButtonUpEvent( const POINT& point );
        void    PaintEvent( HDC hdc );
        void    SizeEvent();

        // Data members

    private:
        // Copy constructor - not allowed
        TreeView( const TreeView& oTreeView );

        // Assignment operator  - not allowed
        TreeView& operator= ( const TreeView& oTreeView );

        // Methods
        void    DrawClientAreaTitleBar( HDC hdc );
        void    DrawItemBackground( HDC hdc, int lineNumber, int yPos, int clientWidth);
        int     DrawItemBitmap( HDC hdc,
                                const TreeViewItem& Item, int xPos, int yPos, bool rtlReading );
        int     DrawItemLabel( HDC hdc,
                               const TreeViewItem& Item, int xPos, int yPos, bool rtlReading );
        int     DrawItemSquare( HDC hDC,
                                const TreeViewItem& Item, int xPos, int yPos, bool rtlReading );
        void    EvaluatePreferredSize( SIZE* pPreferredSize );
  TreeViewItem* GetItemAtLineNumber( int lineNumber, int* pIndex );
  TreeViewItem* GetLastSelectedItem();
        int     GetNumVisibleLines();
        void    HandleLineSelectionEvent( int lineNumber, WPARAM keyState, bool mouseEvent );
        void    InformAppMessageListener();
        int     LineNumberFromIndex( int index );
        int     MouseOverLineNumber( const POINT& point );
        void    SetSelectedItems();
        bool    ShouldShowItem( const TreeViewItem* pItem, BYTE* pCurrentIndent );
        void    UpdateScrollPosition( int lineNumber );

        // Data members
        int          NODE_SQUARE_SIZE;      // pseudo-constant
        int          m_nIdxSelectedFirst;
        int          m_nIdxSelectedLast;
        int          m_nHiLiteLine;
        WORD         m_uCollapseAllBitmapId;
        WORD         m_uExpandAllBitmapId;
        WORD         m_uSelectAllBitmapId;
        WORD         m_uDeselectAllBitmapId;
        COLORREF     m_crClientTitleBar;
        COLORREF     m_crTextHiLiteBackground;
        HBITMAP      m_hBitmapNodeOpen;
        HBITMAP      m_hBitmapNodeClosed;
        HBITMAP      m_hBitmapLeaf;
        WORD         m_wIndentTab;
        RECT         m_LastWindowRect;
        RECT         m_LastHiLiteLineRect;
        String       m_CollapseAll;
        String       m_ExpandAll;
        String       m_SelectAll;
        String       m_DeselectAll;
        String       m_ClientTitleBarText;
        TList<TreeViewItem> m_Items;
};

#endif  // PXSBASE_TREE_VIEW_H_
