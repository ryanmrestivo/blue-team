///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Scroll Pane Class Header
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

#ifndef PXSBASE_SCOLL_PANE_H_
#define PXSBASE_SCOLL_PANE_H_

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

class ScrollPane : public Window
{
    public:
        // Default constructor
        ScrollPane();

        // Destructor
        ~ScrollPane();

        // Methods
        int     GetNumClientLines();
        bool    GetScrollPosition( POINT* pScrollPosition );
        bool    SetScrollPosition( const POINT& scrollPosition );
        void    UpdateScrollBarsInfo( const SIZE& size );

    protected:
        // Methods
        LRESULT KeyDownEvent( WPARAM virtKey );
        void    MouseWheelEvent( WORD keys, SHORT delta, const POINT& point );
        LRESULT ScrollEvent( bool vertical, int scrollCode );

        // Data members
        SHORT   m_sMouseWheelDelta;  // Total rotation since last WHEEL_DELTA
        int     m_nScreenLineHeight;

    private:
        // Copy constructor - not allowed
        ScrollPane( const ScrollPane& oScrollPane );

        // Assignment operator - not allowed
        ScrollPane& operator= ( const ScrollPane& oScrollPane );

        // Methods

        // Data members
};

#endif  // PXSBASE_SCOLL_PANE_H_
