///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Spinner Control Class Header
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

#ifndef PXSBASE_SPINNER_H_
#define PXSBASE_SPINNER_H_

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
#include "PxsBase/Header Files/TextField.h"
#include "PxsBase/Header Files/Window.h"

// 6. Forwards

///////////////////////////////////////////////////////////////////////////////////////////////////
// Interface
///////////////////////////////////////////////////////////////////////////////////////////////////

class Spinner : public Window
{
    public:
        // Default constructor
        Spinner();

        // Destructor
        ~Spinner();

        // Methods
        int     Create( HWND hWndParent );
        void    GetRange( DWORD* pLow, DWORD* pHigh );
        DWORD   GetValue();
        void    SetEnabled( bool enabled );
        void    SetRange( DWORD low, DWORD high );
        void    SetValue( DWORD value );

    protected:
        // Methods
        LRESULT CommandEvent( WPARAM wParam, LPARAM lParam );
        void    GainFocusEvent();
        void    MouseLButtonDblClickEvent( const POINT& point );
        void    MouseLButtonDownEvent( const POINT& point, WPARAM keys );
        void    MouseLButtonUpEvent( const POINT& point );
        void    PaintEvent( HDC hdc );
        void    TimerEvent( UINT_PTR timerID );

        // Data members
        const UINT_PTR  TIMER_ID;
        const DWORD     TIMER_INTERVAL;

    private:
        // Copy constructor - not allowed
        Spinner( const Spinner& oSpinner );

        // Assignment operator - not allowed
        Spinner& operator= ( const Spinner& oSpinner );

        // Methods
        void    HandleValueChangeEvent( const POINT& point );

        // Data members
        int         WIDTH_BUTTONS;
        bool        m_bMouseWasPressed;
        bool        m_bFirstTimerEvent;
        DWORD       m_uHigh;
        DWORD       m_uLow;
        DWORD       m_uValue;
        TextField   m_TextField;
};

#endif  // PXSBASE_SPINNER_H_
