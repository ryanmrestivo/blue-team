///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Image Button Class Header
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

#ifndef PXSBASE_IMAGE_BUTTON_H_
#define PXSBASE_IMAGE_BUTTON_H_

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
#include "PxsBase/Header Files/StaticControl.h"
#include "PxsBase/Header Files/Window.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class ImageButton : public Window
{
    public:
        // Default constructor
        ImageButton( int buttonWidth  = 64,
                     int buttonHeight = 52,
                     int imageWidth   = 32,
                     int imageHeight  = 32 );

        // Destructor
        ~ImageButton();

        // Methods
        void  GetImageSize( SIZE* pSize ) const;
        void  GetPreferredSize( SIZE* pSize );
const String& GetText() const;
        bool  IsCaptionDisplayed() const;
        void  SetBitmap( WORD resourceID, COLORREF transparent );
        void  SetDisplayCaption( bool displayCaption );
        void  SetEnabled( bool enabled );
        void  SetFilledBitmap( COLORREF fill, COLORREF border, DWORD shape );
        void  SetIcons( WORD defaultID, WORD onID, WORD offID );
        void  SetText( const String& Text );

    protected:
        // Methods
        void    MouseLButtonDownEvent( const POINT& point, WPARAM keys );
        void    MouseLButtonUpEvent( const POINT& point );
        void    MouseMoveEvent( const POINT& point, WPARAM keys );
        void    MouseRButtonDownEvent( const POINT& point );
        void    PaintEvent( HDC hdc );
        void    TimerEvent( UINT_PTR timerID );

        // Data members

    private:
        // Copy constructor - not allowed
        ImageButton( const ImageButton& oImageButton );

        // Assignment operator - not allowed
        ImageButton& operator= ( const ImageButton& oImageButton );

        // Data members
        bool          m_bMouseInside;      // Indicates mouse is inside
        bool          m_bMouseWasPressed;  // Left button pressed inside button
        bool          m_bDisplayCaption;
        UINT          m_uTimerID;
        SIZE          m_ImageSize;
        HICON         m_hIconOn;
        HICON         m_hIconOff;
        HICON         m_hIconDefault;
        HBITMAP       m_hBitmapOn;
        HBITMAP       m_hBitmapOff;
        HBITMAP       m_hBitmapDefault;
        StaticControl m_StaticImage;
        StaticControl m_StaticText;

        // Methods
        void    MakeBitmaps( HBITMAP hBitmapDefault, COLORREF transparent );
};

#endif  // PXSBASE_IMAGE_BUTTON_H_
