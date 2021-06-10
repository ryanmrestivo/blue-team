///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Static Drawable Control Class Header
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

#ifndef PXSBASE_STATIC_CONTROL_H_
#define PXSBASE_STATIC_CONTROL_H_

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
#include "PxsBase/Header Files/StringT.h"
#include "PxsBase/Header Files/Font.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class StaticControl
{
    public:
        // Default constructor
        StaticControl();

        // Copy constructor
        StaticControl( const StaticControl& oStaticControl );

        // Destructor
        ~StaticControl();

        // Assignment operators
        StaticControl& operator= ( const StaticControl& oStaticControl );

        // Methods
        void        Draw( HDC hdc );
        DWORD       GetAlignmentX() const;
        DWORD       GetAlignmentY() const;
        COLORREF    GetBackground() const;
        void        GetBounds( RECT* pBounds ) const;
    const Font&     GetFont() const;
        COLORREF    GetForeground() const;
        int         GetPadding() const;
        DWORD       GetShape() const;
    const String&   GetText() const;
        bool        GetVisible() const;
        bool        IsEnabled() const;
        bool        IsHyperlink() const;
        void        Reset();
        void        SetAlignmentX( DWORD alignmentX );
        void        SetAlignmentY( DWORD alignmentY );
        void        SetBackground( COLORREF background );
        void        SetBackgroundGradient( COLORREF colour1,
                                           COLORREF colour2, bool verticalGradient );
        void        SetBitmap( HBITMAP hBitmap );
        void        SetBounds( const RECT& bounds );
        void        SetEnabled( bool enabled );
        void        SetFont( const Font& Font );
        void        SetForeground( COLORREF foreground );
        void        SetHyperlink( bool hyperlink );
        void        SetIcon( HICON icon );
        void        SetLocation( long x, long y );
        void        SetPadding( int padding);
        void        SetShape( DWORD shape );
        void        SetShapeColour( COLORREF colour );
        void        SetSingleLineText( const String& SingleLineText );
        void        SetText( const String& Text );
        void        SetVisible( bool visible );

    protected:
        // Methods

        // Data members

    private:
        // Methods
        void        DrawShape( HDC hdc );
        void        DrawStaticBackground( HDC hdc );
        void        DrawStaticBitmap( HDC hdc );
        void        DrawStaticIcon( HDC hdc );
        void        DrawStaticShape( HDC hdc );
        void        DrawStaticText( HDC hdc );
        UINT        GetDrawTextFormat();
        COLORREF    GetTextColour();
        bool        IsClosedShape();

        // Data members
        bool        m_bEnabled;
        bool        m_bVerticalGradient;
        bool        m_bHyperlink;
        bool        m_bVisible;
        bool        m_bSingleLine;
        int         m_nPadding;
        DWORD       m_uShape;
        DWORD       m_uAlignmentX;
        DWORD       m_uAlignmentY;
        RECT        m_rBounds;
        COLORREF    m_crBackground;
        COLORREF    m_crForeground;
        COLORREF    m_crGradient1;
        COLORREF    m_crGradient2;
        COLORREF    m_crShapeColour;
        Font        m_Font;
        String      m_Text;
        HICON       m_hIcon;
        HBITMAP     m_hBitmap;
};

#endif  // PXSBASE_STATIC_CONTROL_H_
