///////////////////////////////////////////////////////////////////////////////////////////////////
//
// Tool Tip Class Header
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

#ifndef PXSBASE_TOOL_TIP_H_
#define PXSBASE_TOOL_TIP_H_

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

class ToolTip : public Window
{
    public:
        // Default constructor  - private
        // Copy constructor     - not allowed
        // Destructor           - private
        // Assignment operator  - not allowed

        // Static methods
        static void         DestroyInstance();
        static ToolTip*     GetInstance();
        static bool         IsInstanceCreated();

        // Methods
const String& GetText() const;
        void  HideTip();
        void  MouseButtonDown( const POINT& point );
        void  SetDismissDelay( DWORD dismissDelay );
        void  SetInitialDelay( DWORD initialDelay );
        void  SetText( const String& TipText );
        void  Show( const String& TipText, const RECT& owner );

    protected:
        // Methods
        void    PaintEvent( HDC hdc );
        void    TimerEvent( UINT_PTR timerID );

        // Data members

    private:
        // Default constructor, private as Singleton class
        ToolTip();

        // Copy constructor - not allowed
        ToolTip( const ToolTip& oToolTip );

        // Destructor, private as Singleton class
        ~ToolTip();

        // Assignment operator - not allowed
        ToolTip& operator= ( const ToolTip& oToolTip );

        // Methods
        void    Init();

        // Data members
        bool            m_bMouseInside;
        bool            m_bMousePressed;
        DWORD           m_uInitialDelay;  // Milliseconds
        DWORD           m_uDismissDelay;  // Milliseconds
        DWORD           m_uElapsedTime;   // Milliseconds since timer started
        RECT            m_OwnerBounds;    // Rectangle that owns the tool tip
        const UINT_PTR  TIMER_ID;
        const DWORD     TIMER_INTERVAL;
        static ToolTip* m_pInstance;      // Pointer to instance of this class
        StaticControl   m_Tip;
};

#endif  // PXSBASE_TOOL_TIP_H_
