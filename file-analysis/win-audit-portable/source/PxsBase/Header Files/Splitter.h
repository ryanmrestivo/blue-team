///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Splitter Class Header
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

#ifndef PXSBASE_SPLITTER_H_
#define PXSBASE_SPLITTER_H_

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
#include "PxsBase/Header Files/Window.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class Splitter : public Window
{
    public:
        // Default constructor
        Splitter();

        // Destructor
        ~Splitter();

        // Methods
        void    AttachWindows( const Window* pOne, const Window* pTwo );
        void    DoLayout();
        int     GetOffset() const;
        int     GetOffsetTwo() const;
        void    SetOffset( int offset );
        void    SetOffsetTwo( int offsetTwo );

    protected:
        // Methods
        void    MouseLButtonDownEvent( const POINT& point, WPARAM keys );
        void    MouseLButtonUpEvent( const POINT& point );
        void    MouseMoveEvent( const POINT& point, WPARAM keys );
        void    PaintEvent( HDC hdc );

        // Methods

    private:
        // Copy constructor - not allowed
        Splitter( const Splitter& oSplitter );

        // Assignment operator - not allowed
        Splitter& operator= ( const Splitter& oSplitter );

        // Methods

        // Data members
        const int DIVIDER_THICKNESS;
        bool      m_bPanelsAttached;
        bool      m_bMouseWasPressed;
        HWND      m_hWndOne;
        HWND      m_hWndTwo;
        POINT     m_SplitterLocation;
        POINT     m_InitialCursorLocation;
};

#endif  // PXSBASE_SPLITTER_H_
