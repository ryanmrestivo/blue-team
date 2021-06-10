///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Tab Window Class Header
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

#ifndef PXSBASE_TAB_WINDOW_H_
#define PXSBASE_TAB_WINDOW_H_

///////////////////////////////////////////////////////////////////////////////////////////////////
// Includes
///////////////////////////////////////////////////////////////////////////////////////////////////

// 1. Own Interface
#include "PxsBase/Header Files/PxsBase.h"

// 2. C System Files

// 3. C++ System Files

// 4. Other Libraries

// 5. This Project
#include "PxsBase/Header Files/Window.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class TabWindow : public Window
{
    public:
        // Default constructor
        TabWindow();

        // Destructor
        ~TabWindow();

        // Methods
        size_t  Add( HBITMAP hBitmap,
                     LPCWSTR pszName,
                     LPCWSTR pszToolTip, bool allowClose, DWORD tabID, Window* pWindow );
        void    DoLayout();
        int     GetMaxTabWidth() const;
        size_t  GetNumberTabs() const;
        size_t  GetSelectedTabIndex() const;
        void    GetSelectedTabName( String* pSelectedTabName ) const;
        Window* GetSelectedTabWindow() const;
        void    GetTabBounds( RECT* pBounds );
        DWORD   GetTabIdentifier( size_t tabIndex ) const;
        DWORD   GetTabState( size_t tabIndex ) const;
        Window* GetWindow( size_t tabIndex ) const;
        Window* GetWindowFromIdentifier( DWORD tabIdentifier ) const;
        int     GetPreferredTabStripHeight() const;
        void    RemoveTab( size_t tabIndex );
        void    SetCloseBitmaps( WORD closeID, WORD closeOnID );
        void    SetMaxTabWidth( int maxTabWidth );
        void    SetNameAndToolTip( size_t tabIndex, const String& Name, const String& ToolTip );
        void    SetSelectedTabIndex( size_t selectedTabIndex );
        void    SetTabColours( COLORREF tabGradient1,
                               COLORREF tabGradient2, COLORREF tabHiLiteBackground );
        void    SetTabBitmap( size_t tabIndex, WORD resourceID );
        void    SetTabState( size_t tabIndex, DWORD state );
        void    SetTabStripHeight( int tabStripHeight );
        void    SetTabVisible( size_t tabIndex, bool visible );
        void    SetTabWidth( int tabWidth );

    protected:
        typedef struct  TYPE_TAB_DATA
        {
            BOOL    visible;
            BOOL    allowClose;
            DWORD   tabID;
            DWORD   state;
            RECT    bounds;
            HBITMAP hBitmap;
            Window* pWindow;
            wchar_t szToolTip[ MAX_PATH ];
            wchar_t szName[ MAX_PATH ];
        }
        TYPE_TAB_DATA, *PTYPE_TAB_DATA;

        // Data members
        int      HORIZ_TAB_GAP;         // pseudo-constant
        int      INTERNAL_LEADING;      // pseudo-constant
        int      BASELINE_HEIGHT;       // pseudo-constant
        int      MIN_TAB_WIDTH;         // pseudo-constant
        int      MIN_TAB_HEIGHT;        // pseudo-constant

        // Methods
        void    MouseLButtonDownEvent( const POINT& point, WPARAM keys );
        void    MouseLButtonUpEvent( const POINT& point );
        void    MouseMoveEvent( const POINT& point, WPARAM keys );
        void    MouseRButtonDownEvent( const POINT& point );
        void    NotifyEvent( const NMHDR* pNmhdr );
        void    PaintEvent( HDC hDC );
        void    TimerEvent( UINT_PTR timerID );

    private:
        // Copy constructor - not allowed
        TabWindow( const TabWindow& oTabWindow );

        // Assignment operator - not allowed
        TabWindow& operator= ( const TabWindow& oTabWindow );

        // Methods
        void    DrawTabBitmapAndText( HDC hDC, const TYPE_TAB_DATA& TabData );
        void    DrawTabCloseBitmap( HDC hDC, const RECT& tabBounds );
        void    DrawTabShape( HDC hDC, size_t tabIndex );
        int     EvaluateTabWidth( HDC hDC, int tabHeight, const TYPE_TAB_DATA& TabData );
        void    LayoutPages( const RECT& bounds );
        size_t  MouseInTabIndex( const POINT& point );

        // Data members
        int      m_nPadding;
        int      m_nTabWidth;
        int      m_nMaxTabWidth;
        int      m_nTabStripHeight;
        UINT     m_uTimerID;
        size_t   m_uSelectedTabIndex;
        size_t   m_uMouseInTabIndex;
        HBITMAP  m_hBitmapClose;
        HBITMAP  m_hBitmapCloseOn;
        COLORREF m_crTabGradient1;
        COLORREF m_crTabGradient2;
        COLORREF m_crTabHiLiteBackground;
        TArray< TYPE_TAB_DATA> m_Tabs;
};

#endif  // PXSBASE_TAB_WINDOW_H_
