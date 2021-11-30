///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Window Class Header
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

#ifndef PXSBASE_WINDOW_H_
#define PXSBASE_WINDOW_H_

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
#include "PxsBase/Header Files/Font.h"
#include "PxsBase/Header Files/Library.h"
#include "PxsBase/Header Files/StaticControl.h"
#include "PxsBase/Header Files/TArray.h"

// 6. Forwards
class ImageButton;
class MenuItem;
class String;

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class Window
{
    public:
        // Default constructor
        Window();

        // Destructor
        virtual ~Window();

        // Methods
        void     AppendText( const String& Text );
        int      Create( HWND hWndParent );
        int      Create( const Window* pParentWindow );
        void     DestroyWindowHandle();
    virtual void DoLayout();
        void     DoPaintEvents( UINT_PTR timerIDEvent );
        void     Focus();
        UINT     GetAppMessageCode() const;
        HWND     GetAppMessageListener() const;
        COLORREF GetBackground() const;
        DWORD    GetBorderStyle() const;
        void     GetBounds( RECT* pBounds) const;
        void     GetClientSize( SIZE* pClientSize ) const;
        bool     GetDoubleBuffered() const;
        DWORD    GetExStyle() const;
    const Font&  GetFont() const;
        COLORREF GetForeground() const;
        HWND     GetHwnd() const;
        void     GetLayoutProperties( DWORD* pLayoutStyle, int* pHorizontalGap, int* pVerticalGap);
        void     GetLocation( POINT* pLocation ) const;
        void     GetSize( SIZE* pSize ) const;
        LONG     GetStyle() const;
const String&    GetToolTipText() const;
        void     GetText( String* pText ) const;
        void     GetWndClassExClassName( String* pClassName ) const;
        bool     IsEnabled() const;
        bool     IsRightToLeftReading();
        bool     IsVisible() const;
        void     RectangleToScreen( RECT* pBounds ) const;
        void     RedrawAllNow() const;
        void     Repaint() const;
        void     RepaintChildren() const;
        void     RtlMirror( int parentWidth );
        void     RtlStatics();
        void     ScrollToBottom();
        void     ScrollToTop();
        void     SetAppMessageCode( UINT appMessageCode );
        void     SetAppMessageListener( HWND hWndAppMessageListener );
        void     SetBackground( COLORREF background );
        void     SetBackgroundGradient( COLORREF gradient1,
                                        COLORREF gradient2, bool gradientVertical );
        void     SetBounds( const RECT& bounds );
        void     SetBounds( int x, int y, int width, int height );
        void     SetBorderStyle( DWORD borderStyle );
        void     SetHideBitmaps( HWND hideListener, WORD hideID, WORD hideOnID );
        void     SetClipboardRichText( const String& RichText );
        void     SetClipboardText( const String& Text );
        void     SetColours( COLORREF background, COLORREF foreground, COLORREF highlight );
        void     SetDoubleBuffered( bool doubleBuffered );
        void     SetEnabled( bool enabled );
        void     SetForeground( COLORREF foreground );
        void     SetExStyle( DWORD exStyle, bool add );
        void     SetFont( const Font& WindowFont );
        void     SetLayoutProperties( DWORD layoutStyle, int horizontalGap, int verticalGap );
        void     SetLocation( int x, int y );
        void     SetLocation( const POINT& position );
        void     SetRightToLeftReading( bool rightToLeftReading );
        void     SetSize( int width, int height );
        void     SetSize( const SIZE& size );
        void     SetStyle( LONG style, bool add );
        void     SetText( const String& Text );
        void     SetToolTipText( const String& ToolTipText );
        void     SetToTopZ();
        void     SetVisible( bool visible );

    protected:
        // Methods
        void    DrawBackground( HDC hdc );
        void    DrawHideBitmap( HDC hDC );
        void    PositionHideBitmap();

        // Event Handlers
        virtual void    ActivateEvent( bool activated );
        virtual void    AppMessageEvent( UINT uMsg, WPARAM wParam, LPARAM lParam );
        virtual LRESULT ColorStaticEvent( HWND hWndStatic );
        virtual void    CloseEvent();
        virtual LRESULT CommandEvent( WPARAM wParam, LPARAM lParam );
        virtual void    CopyDataEvent( HWND hWnd, const COPYDATASTRUCT* pCds );
        virtual void    DestroyEvent();
        virtual void    DisplayChangeEvent();
        virtual void    DrawMenuItemEvent( const DRAWITEMSTRUCT* pDraw );
        virtual LRESULT EraseBackgroundEvent( HDC hdc );
        virtual void    GainFocusEvent();
        virtual UINT    GetDlgCodeEvent();
        virtual void    HelpEvent( const HELPINFO* pHelpInfo );
        virtual LRESULT KeyDownEvent( WPARAM virtKey );
        virtual void    LayoutChildren();
        virtual void    LostFocusEvent( HWND hWnd );
        virtual void    MeasureMenuItemEvent( MEASUREITEMSTRUCT* pMeasure );
        virtual LRESULT MenuCharEvent( wchar_t chUser , HMENU hMenu );
        virtual void    MouseLButtonDblClickEvent( const POINT& point );
        virtual void    MouseLButtonDownEvent(const POINT& point, WPARAM keys);
        virtual void    MouseLButtonUpEvent( const POINT& point );
        virtual void    MouseMoveEvent( const POINT& point, WPARAM keys );
        virtual void    MouseRButtonDownEvent( const POINT& point );
        virtual void    MouseRButtonUpEvent( const POINT& point );
        virtual void    MouseWheelEvent( WORD keys, SHORT delta, const POINT& point );
        virtual void    NonClientPaintEvent( HDC hdc );
        virtual void    NotifyEvent( const NMHDR* pNmhdr );
        virtual void    PaintEvent( HDC hdc );
        virtual LRESULT ScrollEvent( bool vertical, int scrollCode );
        virtual void    SettingChangeEvent();
        virtual void    SizeEvent();
        virtual void    TimerEvent( UINT_PTR timerID );
        static LRESULT  CALLBACK WindowProc( HWND hWnd, UINT Msg, WPARAM wParam, LPARAM lParam );

        // Data members
        bool            m_bGradientVertical;
        bool            m_bUseParentRectGradient;
        bool            m_bControlsCreated;
        UINT            m_uAppMessageCode;
        DWORD           m_uBorderStyle;
        HWND            m_hWindow;
        HWND            m_hWndAppMessageListener;
        HBRUSH          m_hBackgroundBrush;
        COLORREF        m_crBackground;
        COLORREF        m_crBorder;
        COLORREF        m_crBorderHiLight;
        COLORREF        m_crForeground;
        COLORREF        m_crGradient1;
        COLORREF        m_crGradient2;
        COLORREF        m_crMenu;
        COLORREF        m_crMenuHiLight;
        UINT            HIDE_WINDOW_TIMER_ID;        // Psuedo-constant
        RECT            m_recHideBitmap;
        HWND            m_hWndHideListener;
        HBITMAP         m_hBitmapHide;
        HBITMAP         m_hBitmapHideOn;
        WNDCLASSEX      m_WndClassEx;
        CREATESTRUCT    m_CreateStruct;
        Font            m_Font;
        Library         m_RichEditLibrary;
        TArray< StaticControl > m_Statics;

    private:
        // Copy constructor - not allowed
        Window( const Window& oWindow );

        // Assignment operator - not allowed
        Window& operator= ( const Window& oWindow );

        // Methods
        void     CopyDataToClipboard( UINT format, const void* pData, size_t numBytes );
        void     DrawMenuItemBackGround( const DRAWITEMSTRUCT* pDraw, bool menuBarItem );
        void     DrawMenuItemText( const DRAWITEMSTRUCT* pDraw, const MenuItem* pMenuItem );
        void     DrawPopupMenuImageArea( const DRAWITEMSTRUCT* pDraw, HBITMAP hBitmap );
        void     DrawPopupMenuSeparator( const DRAWITEMSTRUCT* pDraw );
        void     DoPaint();
        COLORREF GetMenuItemTextColour( const DRAWITEMSTRUCT* pDraw, const MenuItem* pMenuItem );

        // Data members
        bool        m_bDoubleBuffered;
        int         m_nHorizontalGap;
        int         m_nVerticalGap;
        int         MENU_ITEM_IMAGE_AREA_WIDTH;  // Psuedo-constant
        DWORD       MAX_MENU_ITEM_CHARS;         // Psuedo-constant
        DWORD       m_uLayoutStyle;
        String      m_ToolTipText;
};

#endif  // PXSBASE_WINDOW_H_
